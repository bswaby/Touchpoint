#roles=Finance,FinanceAdmin,ManageTransactions
# -*- coding: utf-8 -*-
"""
Modern Payment Manager System
============================
A complete, responsive payment management system with:
- Program/Division/Payer drill-down navigation
- Integrated payment processing (no external redirects)
- Email history and preview functionality
- Transaction history with filtering
- AJAX-based user experience
- TouchPoint-compatible design

# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub: https://github.com/bswaby/Touchpoint  (40+ free tools)
# ----------------------------------------------------------------
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to
# support continued development, check out:
#
# DisplayCache - church digital signage that integrates with TouchPoint
# https://displaycache.com
#
# TPxi Go - your church contacts, wherever you work.
# Look up anyone in TouchPoint, log calls and emails from Outlook
# or your phone. No tab switching, no lost context.
# https://tpxigo.com
# ----------------------------------------------------------------

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File  
3. Name the Python "ModernPaymentManager" and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Changelog
---------
v2.4.16 - May 2026
  - Fixed: Multi-person family registrations weren't being flagged as
           refund-driven. Example: one parent paid the full reg total
           for two kids; the refund returned the full amount; each
           kid's individual IndDue (half the reg) didn't match the
           refund (full reg), so neither row got the flag. Detection
           now compares the refund amount against the RegId's TOTAL
           IndDue summed across all co-registrants. Both kids' rows
           flag and clear together.
  - Added: Visual connection on the Outstanding row -- when a refund
           covers multiple people, an amber "linked: <name>, <name>"
           chip appears under the refund-issued pill so staff see
           exactly who's tied together.
  - Added: Zero Out button auto-labels with the person count
           ("Zero Out (2 people)") and the confirmation dialog lists
           every co-registrant who will be cleared. Server-side
           process_zero_out_refund now finds all co-registrants on
           the shared RegId and applies an AdjustFee credit to each
           in one call -- so no one is left dangling after the click.

v2.4.15 - May 2026
  - Fixed: SQL error "Column 'TransactionSummary.OrganizationId' is
           invalid in the select list because it is not contained in
           either an aggregate function or the GROUP BY clause"
           caused by v2.4.14's IsRefundDriven subquery referencing
           ts.OrganizationId when the outer query groups by
           o.OrganizationId. Switched the subquery's filter to use
           o.OrganizationId (same value via the LEFT JOIN, but
           visible inside the GROUP BY).

v2.4.14 - May 2026
  - Added: Refund-driven Outstanding rows are now visually flagged
           AND can be cleared with one click. SQL adds an
           IsRefundDriven flag per row -- TRUE when the most recent
           transaction on that (PeopleId, OrgId) is a refund whose
           amount equals the current outstanding (within a cent).
           When TRUE the row renders a red "refund issued" pill under
           the dollar amount and a "Zero Out" button next to it.
  - Added: Zero Out action posts an AdjustFee credit equal to the
           outstanding so the person no longer appears in Outstanding
           Balances. Description is ADJ|Refund cleared: not actually
           due (auto-zeroed), so the Payment History breakdown still
           shows the original payment, the refund, and the new
           clearing adjustment for a complete audit trail. Server
           cross-checks the client-supplied amount against actual
           IndDue before posting -- if anything changed since the
           page loaded (race condition), it bails out with a refresh
           prompt instead of over-crediting.

v2.4.13 - May 2026
  - Added: "Refund" checkbox in the Receipts payment-type filter
           (sixth option, default OFF). Behavior:
             * Refund alone -> all refunds across CC/CHK/CSH/Coupon
             * Refund + Check -> Check payments AND Check refunds
             * Check alone -> only positive-amt Check rows (no refunds)
           Other type checkboxes default to "amt > 0" when Refund is
           unchecked so the receipt list stays pure-payments unless
           you opt in. Refund rows render with a red pill so they're
           hard to miss.

v2.4.12 - May 2026
  - Fixed: Refund-detection split (v2.4.8) was only applied to the
           Receipts list, not to the inline Payment History breakdown.
           Result: a CC payment that was later refunded still showed
           as "Credit Card" instead of "Refund" when you expanded the
           Outstanding row. Added the same amt-sign rule to the
           breakdown's Kind CASE: Response: / CHK / CSH with amt<0
           now classify as "Refund" / "Check Refund" / "Cash Refund".
           Refund pills render in red so staff don't accidentally
           chase a balance the church already returned.

v2.4.11 - May 2026
  - UX:    Outstanding-amount cell now has a "payment history >>"
           sub-label and a clearer tooltip so the inline expand is
           discoverable at a glance. Old "Transactions" button in the
           History column renamed to "All-Org History" with a tooltip
           explaining the scope difference -- the inline caret is the
           per-involvement payment history, the button is the full
           cross-involvement view (separate page).

v2.4.10 - May 2026
  - Added: Post-write verification. After WriteContentPython succeeds,
           we read the saved script back via model.PythonContent and
           confirm the new APP_VERSION marker is present. If not, the
           update was written to a different script name than the one
           currently running.  The error message tells the admin to
           paste the latest code directly into the correctly-named
           script in Admin > Advanced > Special Content > Python.
  - UX:    Success message now names the new version and adds a hard-
           refresh tip (Ctrl/Cmd+Shift+R) for cases where the browser
           caches the old script aggressively.

v2.4.9 - May 2026
  - Fixed: Auto-update's "Update Now" button could fail with a cryptic
           message. Pre-write validation now requires the fetched code
           to contain the unmistakable "TPxi_PaymentManager" +
           "APP_VERSION" markers before writing, so a worker glitch
           returning an HTML error page or JSON envelope can't corrupt
           the running script. Error messages include the direct fetch
           URL so the admin can paste manually if anything fails.
  - UX:    Update failures now use alert() instead of a toast so the
           full hint (with the fetch URL) is readable, not truncated.

v2.4.8 - May 2026
  - Fixed: Refunds were classified as "Credit Card" payments in the
           Receipts list (same Message prefix on the Transaction row).
           Hitting Email on a refund would have sent the payer a
           thank-you receipt for a payment they actually got back.
           Split PaymentType on amt sign: positive amt = the original
           payment kind (Check / Cash / Credit Card), negative amt =
           Refund / Check Refund / Cash Refund. Refund rows show a
           red pill and have the Email button disabled with an
           explanatory tooltip; Print + Customize remain available.
  - Fixed: _build_receipt_html used the legacy get_current_balance()
           which can return 0 spuriously (swallowed exception), so
           receipt emails could show "$0 balance" when the recipient
           actually still owed money. Inlined the bulletproof SUM
           query with defensive parsing -- matches v2.4.7's pattern.
  - Fixed: "Other" payment-type filter switched from amtdue < 0 to
           amt > 0. Manual cash / check / ACH / wire entries with
           amtdue=0 (offset against an implicit charge) now show up
           correctly; previously they were invisible.

v2.4.7 - May 2026
  - Fixed: Inline breakdown could synthesize a bogus "Starting balance"
           row even when the visible Transaction sum already matched the
           actual outstanding. Root cause: the shared get_current_balance()
           helper has an "except: return 0.0" that swallows IronPython
           attribute errors, so when its lookup glitched it returned 0
           -- which our gap calc then treated as the real number, drawing
           a fake negative-balance starter. Now we run an inline SUM
           query, parse the result defensively, and only insert the
           starter when actual_due is reliably greater than running by
           at least a cent.

v2.4.6 - May 2026
  - Fixed: Inline breakdown's running balance was mismatching for
           payers on recurring-billing involvements (e.g. monthly
           setDefaultCharge + Response: CC pairs where the Response
           row has amtdue=0 even though the payment really did pay
           down a prior charge -- summing amtdue would over-report).
           Switched balance math to -amt (cash direction) which works
           across both recurring-billing and refund patterns.
  - Added: Synthetic "Starting balance" row at the top of the breakdown
           whenever the visible Transaction sum doesn't equal the
           actual current balance. Catches implicit registration fees
           that never got written as a separate Transaction row, so
           the running totals land exactly on the outstanding amount.

v2.4.5 - May 2026
  - Fixed: Inline charge breakdown's running balance didn't match the
           Outstanding column. Two root causes:
           1) The "if Message looks like a payment, force negative
              balance impact" override mis-signed REFUNDS (which have a
              payment-shaped Message like 'Response: CC' but a positive
              amtdue because they raise the balance). Switched to trust
              amtdue verbatim as the source of truth -- it's what
              TouchPoint uses internally to compute IndDue/TotDue.
           2) The minus sign in the Payment column was confusing on
              rows with amtdue=0 (online-reg payment + implicit charge
              in the same transaction). Removed the sign and added a
              quiet "(no balance change)" annotation so the staffer can
              see why the running balance didn't move on a net-zero
              transaction.

v2.4.4 - May 2026
  - Added: Adjust Balance modal now has an "Email the payer" checkbox
           (on by default). When checked, after AdjustFee succeeds the
           server sends a self-contained notification email with the
           direction badge (Charge added / Credit applied), amount,
           reason, current balance, and a Pay Now button if a balance
           remains. Uses the program's configured sender; quietly
           skips if no email is on file.
  - UX: Removed the second "Are you sure?" confirm dialog after
           clicking Apply Adjustment. The button already represents the
           commit; the modal preview shows what will happen.

v2.4.3 - May 2026
  - Fixed: Adjust Balance had the sign convention inverted -- Add Charge
           was decreasing the balance and Apply Credit was increasing it.
           Flipped to match TouchPoint's amt-column convention (positive
           amount = payment-shaped / reduces balance; negative = charge-
           shaped / increases balance). UI labels unchanged; only the
           internal sign passed to model.AdjustFee flipped.

v2.4.2 - May 2026
  - Added: Adjust Balance per-row action. New "Adjust" button opens a
           modal with a direction radio (Add Charge / Apply Credit),
           amount, and required reason. Live preview shows the math as
           you type ("Add charge of $50.00 -- Current: $200 -> $250").
           Server calls model.AdjustFee and writes the description as
           ADJ|<reason>, so the inline charge breakdown labels it as an
           Adjustment and the reason text becomes the staff note. Page
           refreshes 700ms after success.

v2.4.1 - May 2026
  - Added: SMS support on Send Pay Link. The single "Send Pay Link"
           button is now two: "Email Link" and "Text Link". Server
           accepts a 'channel' param (email / text / both). Text path
           uses model.SendSms with the first SmsGroups row (matches the
           FMC Member Manager pattern); SMS body includes the
           authenticated paylink + a one-way disclaimer. Success toast
           reports which channel(s) fired.

v2.4.0 - May 2026
  - Fixed: "Error processing payment link: sender_phone" when sending
           a pay link. build_payment_email used bracket access on a
           sender dict that could be missing keys; switched to .get()
           with fallbacks. Same scrub on build_payment_confirmation_
           email's dead-fallback path.
  - Fixed: _pm_normalize_settings now backfills every expected key
           (sender_id, sender_email, sender_alias, email_title,
           sender_phone) on the default sender AND on every per-program
           override -- so a partial Settings save can't leave downstream
           code KeyError-ing on a missing field.

v2.3.9 - May 2026
  - Fixed: Payment Type multi-select dropdown was still being clipped
           even with z-index:10000 because some parent container had
           overflow:hidden. Switched the panel to position:fixed and
           compute its viewport coords from the button's bounding rect
           on open; scroll re-anchors. Now all five options render
           fully regardless of parent CSS.

v2.3.8 - May 2026
  - Fixed: Payment Type dropdown couldn't be closed by clicking
           outside it (native <details> doesn't support that), and the
           panel was clipped by the filter card's rounded corners.
           Rewrote as button + popup div with a document-level click
           listener (closes on outside-click + Esc) and bumped z-index.

v2.3.7 - May 2026
  - UI: Payment type multi-select polish (deployed by user mid-thread
           between fixes; spans 2.3.7-only changes).

v2.3.6 - May 2026
  - Added: Last-4 search field on Receipts. Matches Transaction.LastFourCC
           OR LastFourACH exactly. Card / ACH last-4 chip also surfaces
           inline on every row that has one stored.
  - Added: "Other" payment-type checkbox catches anything that reduced a
           balance without a CHK|/CSH|/Response/Coupon prefix -- manual
           cash/check entries, ACH, wire transfers, scholarship credits,
           refunds. Labelled "Other Payment" in the row pill.
  - Added: Inline charge breakdown on every Outstanding-Balance row.
           Caret expands an inline ledger of charges + payments that
           produced the balance, including running balance, kind chips
           (Adjustment / Credit / Late Fee / Variable Charge / Recurring
           Setup / Transfer / etc.), and the staff note for each row.
           Newest entries top, oldest bottom; running balance computed
           chronologically.
  - Added: Customize Receipt -- per-row modal with a plain-text Personal
           Note field, live preview, alt-email override, and "Print
           Edited" / "Send Customized Receipt" actions. Works on every
           payment type (no PM-origin gating since it generates a fresh
           receipt, not a duplicate of an original send).
  - Added: Inline payer notes on Receipts rows. The sub-line under each
           payer name now shows the staff-typed note (e.g. "Added banquet
           via phone call", "Scholarship approved by Nate McGehee")
           instead of the random TransactionId hex. Coupons still show
           the bare coupon code. Tooltip carries the raw TxnId + full
           note for forensics.
  - Added: Receipt body now includes Involvement + Registrant names + the
           staff note below the chargeNotes line, automatically, with no
           template change required. New {orgName}, {registrants}, and
           {note} merge fields available for templates that want them on
           their own lines.
  - Added: Payer + Involvement substring filters on Receipts.
  - Renamed: top-level nav "Programs" -> "Outstanding Balances" (matches
           accounting language). Internal function names unchanged.
  - Fixed: Payer column now shows the actual cardholder (t.Name) instead
           of the registrant (p.Name). Parents paying for kids now appear
           correctly on the receipt salutation.
  - Fixed: SQL JOIN multiplication on family registrations / multi-use
           coupons. Switched LEFT JOIN to OUTER APPLY (TOP 1) so each
           transaction shows once. Multi-person registrations now list all
           registrants on a single "Registered: ..." line.
  - Fixed: receipts amount was $0 on Credit Card / Coupon rows. Root cause
           was t.amtdue being NULL or 0 for many CC/Coupon transactions.
           Switched to t.amt (consistent with MinistryDepositReport) and
           added t.voided IS NULL guard.
  - Fixed: create_json_response was nesting the data dict under
           response['data'] but every JS handler read fields from the
           root. Silently broke receipts, settings, templates, etc. Merged
           data into root so existing handlers work.
  - Fixed: empty-state diagnostic now distinguishes between "no
           transactions in range", "stale DB snapshot" (surfaces latest
           transaction date), and "wrong message convention".
  - Fixed: caret rotation + toggle on the inline charge breakdown when
           closing/reopening the same row.
  - UI: Programs nav button styled primary (it's the home view); top-bar
           "Go Back" removed from the landing page; "Include payments
           not recorded via Payment Manager" now checked by default.

v2.1.1 - May 2026
  - Added: Email template editor inside Settings. The two HTML templates
           (PM_PaymentMade_Email / PM_Notification_Email) can be edited,
           saved, previewed, and reset to the built-in default without
           leaving the tool. Falls back to defaults so the editor never
           opens empty even before any custom HtmlContent exists.
  - Added: Source column on Receipts results. Distinguishes payments
           recorded via Payment Manager (CHK|/CSH| signature) from those
           recorded elsewhere (TouchPoint batch entry, syncs, etc.).
  - Added: "Include payments not recorded via Payment Manager" toggle
           on Receipts. Default is OFF so a duplicate-receipt email
           isn't fired for a payment that had no original receipt.
  - Defense: Server now refuses to email a duplicate receipt for a
           non-PM payment even if a hand-crafted POST hits the action,
           with a clear explanation in the response.
v2.1.0 - May 2026
  - Added: Receipts tab. Find past check/cash payments by date range
           (single date or range), then Print (popup-window receipt
           for church files) or Email (re-send the original-template
           receipt with a "Duplicate Receipt" banner).
  - Added: Settings UI. All hardcoded module constants (default email
           sender, template names, subject prefix, payment-type prefixes)
           now configurable from a Settings modal. Per-program sender
           overrides replace the legacy PM_EmailSenders dict (auto-
           imported on first save). Stored as JSON in
           TextContent('PaymentManager_Settings').
  - Added: AutoUpdate banner -- one-click pull of newer published
           versions; saved Settings preserved across updates.
  - Added: ? Help modal documenting drill-down flow, Receipts, Settings,
           and AutoUpdate.
  - Internal: get_email_details() now routes through pm_load_settings()
           rather than parsing PM_EmailSenders directly.
v2.0 - May 2025
  - Initial modern rewrite (drill-down navigation, AJAX, modal payment
    record, email history + preview, transaction history).
"""

from datetime import datetime
import sys
import ast
import re
import json
from collections import defaultdict

# =====================================================================
# VERSION / AUTO-UPDATE  (matches the TPxi house pattern)
# =====================================================================
APP_VERSION = '2.4.16'
DC_SCRIPT_ID = 'TPxi_PaymentManager'
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'

def pm_get_script_name():
    """Detect the installed script name (admin may have renamed it)."""
    try:
        if hasattr(model.Data, 'script_name') and model.Data.script_name:
            sn = str(model.Data.script_name).strip()
            if sn:
                return sn
    except:
        pass
    try:
        url = str(getattr(model, 'URL', '') or '')
        m = re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    return DC_SCRIPT_ID

# =====================================================================
# SETTINGS STORAGE
# Single source of truth for what used to be hardcoded module constants.
# Stored as JSON in TextContent('PaymentManager_Settings'); first load
# bootstraps from the legacy PM_EmailSenders dict so existing installs
# don't lose their per-program sender mappings.
# =====================================================================
PM_SETTINGS_KEY = 'PaymentManager_Settings'

# General Settings (defaults -- override via Settings UI)
DEFAULT_PROGRAM_ID = 0
PAGE_TITLE = "Payment Manager"

# Payment Settings
PAYMENT_LINK_DESCRIPTION = "CreditCardLinkSent"
CREDIT_CHARGE_PAYMENT_TYPE = "CreditCharge"
CHECK_PAYMENT_TYPE = "CHK|"
CASH_PAYMENT_TYPE = "CSH|" 

# Email Settings
EMAIL_SUBJECT_PREFIX = "Payment Notification - "
DEFAULT_EMAIL_SENDER = {
    'sender_id': '3134',
    'sender_email': 'noreply@church.com',
    'sender_alias': 'Payment System',
    'email_title': 'Payment Notification',
    'sender_phone': '(555) 123-4567'
}

# Template Settings >>> See "EMAIL TEMPLATE SETUP INSTRUCTIONS" below for template usage <<<
PAYMENT_CONFIRMATION_TEMPLATE_NAME = "PM_PaymentMade_Email"
PAYMENT_NOTIFICATION_TEMPLATE_NAME = "PM_Notification_Email"

# Email Templates -- friendlier defaults that look acceptable out of the
# box. These render when no custom HtmlContent exists for the configured
# template name, OR when Settings -> Email Template Bodies -> Reset is
# clicked. They use inline styles only (no <style> blocks) so they
# survive whatever sanitization email clients throw at them.
DEFAULT_PAYMENT_EMAIL_TEMPLATE = """
<div style="font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#333;max-width:600px;">
  <p>Hi {name},</p>
  <p>You have a balance due. Use the secure link below to pay online whenever you're ready -- the link is unique to your account and is good until your balance is paid.</p>
  <div style="background:#f7fafc;border:1px solid #e2e8f0;border-radius:6px;padding:14px 18px;margin:14px 0;">
    <div style="margin-bottom:4px;">Previous Balance: <strong>{previousDue}</strong></div>
    <div style="margin-bottom:4px;">New Charges: <strong>{chargeNotes}</strong></div>
    <div style="font-size:1.1em;color:#1f4e79;margin-top:6px;">Total Due: <strong>{totalDue}</strong></div>
  </div>
  <p style="text-align:center;margin:20px 0;">
    <a href="{paylink}" style="display:inline-block;background:#1f4e79;color:#fff;text-decoration:none;padding:10px 22px;border-radius:6px;font-weight:600;">Pay Now &raquo;</a>
  </p>
  <p>If the button doesn't work, copy and paste this link into your browser:<br>
    <span style="word-break:break-all;color:#0078d4;font-size:12px;">{paylink}</span></p>
  <p style="margin-top:24px;">Thank you,<br>
    <strong>{sender_alias}</strong><br>
    {sender_phone}<br>
    <a href="mailto:{sender_email}" style="color:#1f4e79;">{sender_email}</a></p>
</div>
"""

DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE = """
<div style="font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#333;max-width:600px;">
  <p>Hi {name},</p>
  <p>We've received your payment -- thank you! Here's a summary for your records:</p>
  <div style="background:#f0f9ff;border:1px solid #cfe3ff;border-radius:6px;padding:14px 18px;margin:14px 0;">
    <div style="margin-bottom:4px;">Payment: <strong>{chargeNotes}</strong></div>
    <div style="margin-bottom:4px;">Previous Balance: <strong>{previousDue}</strong></div>
    <div style="font-size:1.1em;color:#107c10;margin-top:6px;">New Balance: <strong>{newTotalDue}</strong></div>
  </div>
  <p>If you have questions about this payment, reply to this email or call the number below.</p>
  <p style="margin-top:24px;">Thank you,<br>
    <strong>{sender_alias}</strong><br>
    {sender_phone}<br>
    <a href="mailto:{sender_email}" style="color:#1f4e79;">{sender_email}</a></p>
</div>
"""

"""
-----------------------------------
EMAIL TEMPLATE SETUP INSTRUCTIONS
-----------------------------------

This payment manager uses two customizable email templates:

1. PAYMENT_CONFIRMATION_TEMPLATE_NAME (default: "PM_PaymentMade_Email")
   - Used when recording a payment
   - Contains variables:
     - {name} - Payer's name
     - {chargeNotes} - Payment details (type, amount, description)
     - {previousDue} - Previous balance amount with formatting
     - {newTotalDue} - New balance after payment with formatting
     - {sender_alias} - Name of the email sender
     - {sender_phone} - Phone number for contact
     - {sender_email} - Email address for contact

2. PAYMENT_NOTIFICATION_TEMPLATE_NAME (default: "PM_Notification_Email")
   - Used when sending payment links
   - Contains variables:
     - {name} - Payer's name
     - {chargeNotes} - Charge details with formatting
     - {previousDue} - Previous balance amount with formatting
     - {totalDue} - Total amount due with formatting
     - {paylink} - Generated payment link URL
     - {sender_alias} - Name of the email sender
     - {sender_phone} - Phone number for contact
     - {sender_email} - Email address for contact

To set up these templates:
1. Go to Special Content > HTML in TouchPoint
2. Create a new HTML file with the name specified in the configuration
3. Design your email using the variables listed above
4. Save the template

Default templates will be used if the custom templates are not found.
"""

# =====================================================================
# SETTINGS LOAD / SAVE
# =====================================================================
def _pm_normalize_settings(s):
    # Treat empty strings as "missing" so a prior partial save that
    # wrote '' for template names doesn't leave HtmlContent lookups
    # hitting nothing.
    def _or_default(key, default):
        v = s.get(key)
        if v is None or (isinstance(v, basestring) and not v.strip()):
            s[key] = default
    _or_default('pageTitle', PAGE_TITLE)
    _or_default('paymentLinkDescription', PAYMENT_LINK_DESCRIPTION)
    _or_default('creditChargePaymentType', CREDIT_CHARGE_PAYMENT_TYPE)
    _or_default('checkPaymentType', CHECK_PAYMENT_TYPE)
    _or_default('cashPaymentType', CASH_PAYMENT_TYPE)
    _or_default('emailSubjectPrefix', EMAIL_SUBJECT_PREFIX)
    _or_default('paymentConfirmationTemplate', PAYMENT_CONFIRMATION_TEMPLATE_NAME)
    _or_default('paymentNotificationTemplate', PAYMENT_NOTIFICATION_TEMPLATE_NAME)
    s.setdefault('defaultProgramId', DEFAULT_PROGRAM_ID)
    if not isinstance(s.get('defaultSender'), dict):
        s['defaultSender'] = dict(DEFAULT_EMAIL_SENDER)
    if not isinstance(s.get('senders'), dict):
        s['senders'] = {}
    # Guarantee every sender dict (default + per-program overrides)
    # carries the full set of expected keys, even if the admin left
    # some fields blank when saving. Bare-bracket access like
    # email_details['sender_phone'] in email-building functions used
    # to KeyError otherwise.
    SENDER_KEYS = ('sender_id', 'sender_email', 'sender_alias', 'email_title', 'sender_phone')
    def _fill_sender(d):
        if not isinstance(d, dict):
            return dict(DEFAULT_EMAIL_SENDER)
        for k in SENDER_KEYS:
            if k not in d or d[k] is None:
                d[k] = DEFAULT_EMAIL_SENDER.get(k, '')
        return d
    s['defaultSender'] = _fill_sender(s['defaultSender'])
    for pid, sender in list(s['senders'].items()):
        s['senders'][pid] = _fill_sender(sender)
    return s

def pm_load_settings():
    """Load PM settings JSON. Falls back to module constants and -- on
    first run -- imports the legacy PM_EmailSenders dict so installs
    upgrading from <2.1 don't lose sender mappings."""
    try:
        raw = model.TextContent(PM_SETTINGS_KEY)
        if raw:
            data = json.loads(raw)
            if isinstance(data, dict):
                return _pm_normalize_settings(data)
    except:
        pass
    boot = {}
    try:
        legacy = model.TextContent('PM_EmailSenders')
        if legacy:
            parsed = ast.literal_eval(legacy)
            if isinstance(parsed, dict):
                boot['senders'] = {str(k): v for k, v in parsed.items()}
    except:
        pass
    return _pm_normalize_settings(boot)

def pm_save_settings(s):
    try:
        model.WriteContentText(PM_SETTINGS_KEY, json.dumps(s), '')
        return True
    except:
        return False

def pm_setup_issues():
    """Return a list of human-readable problems with the current
    configuration. Empty list means everything is set up. Used by the
    main-page banner to show admins exactly what to fix and which
    features won't work until they do."""
    issues = []
    try:
        s = pm_load_settings()
    except:
        return [("Settings could not be loaded -- click Settings to initialize.",
                 "All email features unavailable")]
    ds = s.get('defaultSender') or {}
    # Default sender must have at least a from-address and sender-id.
    if not (ds.get('sender_email') or '').strip():
        issues.append(("Default sender email is blank.",
                       "Payment-link and confirmation emails will fail."))
    if not str(ds.get('sender_id') or '').strip():
        issues.append(("Default sender PeopleId is blank.",
                       "TouchPoint requires a queued-by PeopleId; sends will be rejected."))
    if (ds.get('sender_email') or '').strip().lower() in ('noreply@church.com', ''):
        # The bundled placeholder -- almost certainly not the church's
        # real address.
        if (ds.get('sender_email') or '').strip().lower() == 'noreply@church.com':
            issues.append(("Default sender email is the bundled placeholder (noreply@church.com).",
                           "Bounces / spam risk; change to a real church address."))
    # Template existence -- check that HtmlContent rows actually exist,
    # so admins know they're rendering the built-in defaults vs their
    # own customized copy.
    for label, key, default_name in [
        ('Payment Confirmation', 'paymentConfirmationTemplate', PAYMENT_CONFIRMATION_TEMPLATE_NAME),
        ('Payment Notification', 'paymentNotificationTemplate', PAYMENT_NOTIFICATION_TEMPLATE_NAME),
    ]:
        nm = (s.get(key) or default_name)
        try:
            body = model.HtmlContent(nm)
        except:
            body = None
        if not body:
            issues.append((
                "Template '" + label + "' (" + nm + ") not found in HTML Content -- using built-in default.",
                "Emails will send, but with the generic default body. Open Settings to customize."
            ))
    return issues

try:
    class ModernPaymentManager:
        """Complete Payment Manager with integrated functionality"""
        
        def __init__(self):
            self.program_id = self.get_program_id()
            self.current_action = self.get_current_action()
            self.division_filter = getattr(model.Data, 'divFilter', '')
            self.search_filters = self.get_search_filters()
            
        def get_program_id(self):
            """Get the program ID from form data"""
            try:
                return str(getattr(model.Data, 'ProgramID', DEFAULT_PROGRAM_ID))
            except:
                return str(DEFAULT_PROGRAM_ID)
                
        def get_current_action(self):
            """Determine current action from URL parameters"""
            try:
                return str(getattr(model.Data, 'action', 'programs'))
            except:
                return 'programs'
                
        def get_search_filters(self):
            """Extract search filters from form data"""
            return {
                'first_name': str(getattr(model.Data, 'FirstNameSearch', '')),
                'last_name': str(getattr(model.Data, 'LastNameSearch', '')),
                'show_all': str(getattr(model.Data, 'ShowAll', ''))
            }
            
        def format_phone(self, phone_number):
            """Format phone number for display"""
            if not phone_number:
                return ""
            phone_str = str(phone_number).strip()
            if len(phone_str) == 10:
                return '(' + phone_str[:3] + ') ' + phone_str[3:6] + '-' + phone_str[6:]
            elif len(phone_str) == 7:
                return '(615) ' + phone_str[:3] + '-' + phone_str[3:]
            return phone_str
            
        def format_currency(self, amount):
            """Format amount as currency"""
            try:
                return '${:,.2f}'.format(float(amount) + 0.00)
            except:
                return '$0.00'
                
        def safe_get_attr(self, obj, attr, default=''):
            """Safely get attribute from object"""
            try:
                return getattr(obj, attr, default)
            except:
                return default
        
        # Data retrieval methods
        def get_programs_with_dues(self):
            """Get all programs that have outstanding dues"""
            sql = """
            SELECT 
                pro.Name AS ProgramName,
                pro.Id AS ProgramId,
                SUM(ts.IndDue) AS Outstanding,
                COUNT(DISTINCT ts.PeopleId) AS PayerCount
            FROM [TransactionSummary] ts
            INNER JOIN [People] p ON ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE 
                ts.IndDue <> 0
                AND ts.IsLatestTransaction = 1
                --AND pro.Id <> 1152
            GROUP BY 
                pro.Name, pro.Id
            ORDER BY pro.Name
            """
            return q.QuerySql(sql)
            
        def get_divisions_with_dues(self, program_id):
            """Get divisions within a program that have outstanding dues"""
            sql = """
            SELECT 
                d.Name AS DivisionName,
                d.Id AS DivisionId,
                o.OrganizationName,
                o.OrganizationId,
                SUM(ts.IndDue) AS Outstanding,
                COUNT(DISTINCT ts.PeopleId) AS PayerCount
            FROM [TransactionSummary] ts
            INNER JOIN [People] p ON ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE 
                ts.IndDue <> 0
                AND ts.IsLatestTransaction = 1
                AND pro.Id = {0}
                {1}
            GROUP BY 
                d.Name, d.Id, o.OrganizationName, o.OrganizationId
            ORDER BY d.Name, o.OrganizationName
            """.format(program_id, self.get_division_filter_sql())
            return q.QuerySql(sql)
            
        def get_payers_with_dues(self, org_id=None, program_id=None):
            """Get individual payers with outstanding dues"""
            where_clause = "ts.IndDue <> 0 AND ts.IsLatestTransaction = 1"
            
            if org_id:
                where_clause += " AND o.OrganizationId = {0}".format(org_id)
            elif program_id:
                where_clause += " AND pro.Id = {0}".format(program_id)
                
            # Add name search filters
            if self.search_filters['first_name']:
                where_clause += " AND p.FirstName LIKE '{0}%'".format(self.search_filters['first_name'])
            if self.search_filters['last_name']:
                where_clause += " AND p.LastName LIKE '{0}%'".format(self.search_filters['last_name'])
                
            sql = """
            SELECT
                pro.Name AS Program,
                pro.Id AS ProgramId,
                d.Name AS Division,
                d.Id AS DivisionId,
                o.OrganizationName,
                o.OrganizationId,
                ts.PeopleId,
                p.Name2,
                p.Age,
                p.FirstName,
                p.LastName,
                p.EmailAddress,
                p.CellPhone,
                p.HomePhone,
                p.FamilyId,
                SUM(ts.TotPaid) AS Paid,
                SUM(ts.TotCoupon) AS Coupons,
                SUM(ts.IndDue) AS Outstanding,
                ts.TranDate,
                -- Refund-driven flag: TRUE when the most recent
                -- Transaction tied to this (PeopleId, OrgId) is a
                -- refund (Response:/CHK/CSH with amt<0) AND the
                -- refunded amount equals the current outstanding to
                -- within a cent. This is the "they paid then we
                -- refunded; now they appear to owe but they don't"
                -- pattern -- staff shouldn't be chasing these.
                -- Refund-driven detection. Compares the refund amount
                -- against the TOTAL IndDue across the whole RegId, not
                -- just this person's slice. Catches both:
                --   * single-person refund (their slice == reg total)
                --   * multi-person family registration where one
                --     payer covered everyone (e.g. one parent paid for
                --     two kids, refund returned the whole amount, now
                --     each kid's row shows $75 outstanding individually
                --     but the RegId total matches the refund).
                (
                    SELECT TOP 1
                        CASE WHEN tlast.amt < 0
                               AND (tlast.[Message] LIKE 'CHK%'
                                    OR tlast.[Message] LIKE 'CSH%'
                                    OR tlast.[Message] LIKE 'Response%')
                               AND ABS(ABS(tlast.amt) - (
                                    SELECT SUM(ISNULL(IndDue, 0))
                                    FROM TransactionSummary
                                    WHERE RegId = tlast.OriginalId AND IsLatestTransaction = 1
                               )) < 0.01
                             THEN 1 ELSE 0 END
                    FROM TransactionSummary tslast
                    INNER JOIN [Transaction] tlast ON tslast.RegId = tlast.OriginalId
                    WHERE tslast.PeopleId = ts.PeopleId
                      AND tslast.OrganizationId = o.OrganizationId
                      AND tslast.IsLatestTransaction = 1
                      AND tlast.amt <> 0
                      AND tlast.voided IS NULL
                    ORDER BY tlast.TransactionDate DESC
                ) AS IsRefundDriven,
                -- Comma-separated list of OTHER people on the same
                -- RegId who also have IndDue > 0. Empty string for
                -- solo registrations. Lets the UI show staff the
                -- visual connection: "this refund affects Tanner +
                -- Sarah Krantz, both will be zeroed together".
                ISNULL((
                    SELECT STUFF((
                        SELECT ', ' + p2.Name
                        FROM TransactionSummary ts2
                        INNER JOIN People p2 ON p2.PeopleId = ts2.PeopleId
                        WHERE ts2.RegId IN (
                            SELECT RegId FROM TransactionSummary
                            WHERE PeopleId = ts.PeopleId
                              AND OrganizationId = o.OrganizationId
                              AND IsLatestTransaction = 1
                        )
                          AND ts2.IsLatestTransaction = 1
                          AND ts2.IndDue > 0
                          AND ts2.PeopleId <> ts.PeopleId
                        FOR XML PATH('')
                    ), 1, 2, '')
                ), '') AS CoRegistrantNames
            FROM [TransactionSummary] ts
            INNER JOIN [People] p ON ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE {0}
            GROUP BY
                d.Name, o.OrganizationName, o.OrganizationId, pro.Name, pro.Id,
                d.Name, d.Id, ts.PeopleId, p.Name2, p.Age, p.FirstName, p.LastName,
                p.EmailAddress, p.CellPhone, p.HomePhone, p.FamilyId, ts.TranDate
            ORDER BY o.OrganizationName, p.Name2
            """.format(where_clause)
            return q.QuerySql(sql)
            
        def get_division_filter_sql(self):
            """Generate SQL filter for division"""
            if self.division_filter and self.division_filter != 'All':
                try:
                    div_id = int(self.division_filter)
                    return " AND d.Id = {0}".format(div_id)
                except:
                    return " AND d.Name = '{0}'".format(self.division_filter)
            return ""
            
        def get_parent_emails(self, family_id):
            """Get parent email addresses for CC purposes"""
            try:
                sql = """
                SELECT DISTINCT p.PeopleId, p.EmailAddress, p.FirstName, p.LastName, 
                       p.CellPhone, p.HomePhone 
                FROM dbo.People AS p 
                INNER JOIN dbo.Families AS t1 ON t1.FamilyId = p.FamilyId 
                WHERE (p.FamilyId = {0}) 
                    AND (p.PositionInFamilyId = 10) 
                    AND (NOT (p.IsDeceased = 1)) 
                    AND (NOT (p.ArchivedFlag = 1)) 
                    AND p.EmailAddress <> '' 
                    AND (t1.HeadOfHouseholdId = p.PeopleId OR t1.HeadOfHouseholdSpouseId = p.PeopleId)
                """.format(family_id)
                
                parents = q.QuerySql(sql)
                emails = []
                parent_info = {}
                
                for parent in parents:
                    email = self.safe_get_attr(parent, 'EmailAddress')
                    if email:
                        emails.append(email)
                        parent_info = {
                            'id': self.safe_get_attr(parent, 'PeopleId'),
                            'email': email,
                            'first_name': self.safe_get_attr(parent, 'FirstName'),
                            'last_name': self.safe_get_attr(parent, 'LastName'),
                            'phone': self.format_phone(self.safe_get_attr(parent, 'CellPhone')) or 
                                   self.format_phone(self.safe_get_attr(parent, 'HomePhone'))
                        }
                        
                    # Ensure user account exists for payment links
                    people_id = self.safe_get_attr(parent, 'PeopleId')
                    if people_id:
                        user_count = q.QuerySqlInt("SELECT COUNT(UserId) FROM Users WHERE PeopleId = {0}".format(people_id))
                        if user_count == 0:
                            model.AddRole(people_id, "Access")
                            model.RemoveRole(people_id, "Access")
                
                return ','.join(emails), parent_info
            except Exception as e:
                return '', {}

        def get_email_history(self, people_id):
            """Get email history for a person"""
            sql = '''
            SELECT Top 50
                p.Name, eq.Subject, eq.Body, p.PeopleId, eq.Sent, eq.FromName, 
                eq.Id AS messageId, 
                Count(er.Id) AS Opened
            FROM dbo.EmailQueueTo eqt
            INNER JOIN dbo.EmailQueue eq ON (eqt.Id = eq.Id) 
            INNER JOIN dbo.People p ON (eqt.PeopleId = p.PeopleId)
            LEFT JOIN EmailResponses er ON er.PeopleId = p.PeopleId AND eq.Id = er.EmailQueueId
            WHERE eqt.PeopleId = {0}
            GROUP BY p.Name, eq.Subject, eq.Body, p.PeopleId, eq.Sent, eq.FromName, eq.Id
            ORDER BY eq.Sent DESC
            '''.format(people_id)
            return q.QuerySql(sql)

        def get_charge_details(self, people_id, org_id):
            """Return the underlying transactions for a single
            (PeopleId, OrgId) pair, oldest first so a running balance
            can be calculated client-side.

            Returns rows with: TranDate, Description, Message, Type,
            Amount (signed: + = charge, - = payment), TransactionId.
            """
            try:
                pid = int(people_id)
                oid = int(org_id)
            except:
                return []
            sql = """
                SELECT
                    FORMAT(t.TransactionDate, 'yyyy-MM-dd HH:mm') AS TranDate,
                    t.TransactionId,
                    ISNULL(t.[Message], '')     AS Message,
                    ISNULL(t.[Description], '') AS Description,
                    t.amt                       AS RawAmt,
                    t.amtdue                    AS AmtDue,
                    -- Refund rows (negative amt with a payment-shaped
                    -- Message) must be classified BEFORE the generic
                    -- payment patterns, otherwise a 'Response: CC' row
                    -- with amt=-75 (the bank reversed a payment) ends
                    -- up labeled "Credit Card" instead of "Refund".
                    CASE WHEN t.[Message] LIKE 'CHK%' AND t.amt < 0      THEN 'Check Refund'
                         WHEN t.[Message] LIKE 'CSH%' AND t.amt < 0      THEN 'Cash Refund'
                         WHEN t.[Message] LIKE 'Response%' AND t.amt < 0 THEN 'Refund'
                         WHEN t.[Message] LIKE 'CHK%'                     THEN 'Check'
                         WHEN t.[Message] LIKE 'CSH%'                     THEN 'Cash'
                         WHEN t.[Message] LIKE 'Response%'                THEN 'Credit Card'
                         WHEN t.TransactionId LIKE 'Coupon%'              THEN 'Coupon'
                         WHEN t.[Message] LIKE 'FEE%'                     THEN 'Fee'
                         WHEN t.[Message] LIKE 'ADJ|%' OR t.[Message] LIKE 'Adj%' THEN 'Adjustment'
                         WHEN t.[Message] LIKE 'variableCredit%' OR t.[Message] LIKE 'variableRefund%' OR t.[Message] LIKE '%Credit for%' THEN 'Credit'
                         WHEN t.[Message] LIKE 'variableLate%'            THEN 'Late Fee'
                         WHEN t.[Message] LIKE 'variable%'                THEN 'Variable Charge'
                         WHEN t.[Message] LIKE 'setDefaultCharge%'        THEN 'Recurring Setup'
                         WHEN t.[Message] LIKE 'move-to-payer%'           THEN 'Transfer'
                         WHEN t.[Message] LIKE 'Initial%'                 THEN 'Initial Charge'
                         WHEN t.[Message] LIKE '%Household%'              THEN 'Household Charge'
                         WHEN t.[Message] IS NULL OR t.[Message] = ''     THEN 'Manual'
                         ELSE 'Charge' END AS Kind
                FROM TransactionSummary ts
                INNER JOIN [Transaction] t ON ts.RegId = t.OriginalId
                WHERE ts.PeopleId = {0}
                  AND ts.OrganizationId = {1}
                  AND ts.IsLatestTransaction = 1
                  AND t.amt <> 0
                  AND t.voided IS NULL
                ORDER BY t.TransactionDate ASC, t.TransactionId
            """.format(pid, oid)
            return q.QuerySql(sql)

        def process_charge_details(self):
            """AJAX endpoint backing the inline expand on payer rows.
            Returns a normalized list of charges + payments with a
            running balance computed server-side so all clients see the
            same math.
            """
            try:
                people_id = str(getattr(model.Data, 'people_id', '') or '').strip()
                org_id    = str(getattr(model.Data, 'org_id', '') or '').strip()
                if not people_id or not org_id:
                    return self.create_json_response(False, "people_id and org_id are required")
                rows = []
                running = 0.0
                for r in self.get_charge_details(people_id, org_id):
                    msg = str(self.safe_get_attr(r, 'Message', '') or '')
                    txid = str(self.safe_get_attr(r, 'TransactionId', '') or '')
                    amtdue_raw = self.safe_get_attr(r, 'AmtDue', None)
                    amt    = float(self.safe_get_attr(r, 'RawAmt', 0) or 0)
                    # Balance math: trust -amt as the balance impact.
                    # Verified against two real patterns in local data:
                    #   * payment+refund pair (Response: with negative
                    #     amt) where the visible rows leave a gap that
                    #     a prior registration fee fills in
                    #   * recurring setDefaultCharge + Response: CC
                    #     pairs where the payment side has amtdue=0 --
                    #     summing amtdue would over-report. Summing
                    #     -amt matches actual outstanding.
                    # Positive amt = money in -> reduces balance.
                    # Negative amt = charge committed or refund -> raises
                    # balance.
                    balance_delta = -amt
                    # Column placement: based on amt sign (the *cash*
                    # direction). Positive amt = Payment column.
                    # Negative amt = Charge column. amt=0 is filtered
                    # out at the SQL level.
                    is_payment_row = (amt > 0)
                    display_abs = abs(amt)
                    running += balance_delta
                    # Strip the type prefix (CHK|, CSH|, FEE|, ADJ|)
                    # from the Message to expose just the staff note.
                    note = ''
                    if '|' in msg:
                        note = msg.split('|', 1)[1].strip()
                    elif msg.startswith('Response'):
                        note = ''
                    else:
                        note = msg
                    desc_from_tx = str(self.safe_get_attr(r, 'Description', '') or '').strip()
                    rows.append({
                        'tranDate':    self.safe_get_attr(r, 'TranDate', ''),
                        'transactionId': txid,
                        'kind':        self.safe_get_attr(r, 'Kind', ''),
                        'description': desc_from_tx,
                        'note':        note,
                        'message':     msg,
                        'amount':      display_abs,
                        'balanceDelta': balance_delta,
                        'isPayment':   is_payment_row,
                        'runningBalance': running,
                    })
                # Reconcile against actual current balance. Inline +
                # bulletproofed because the legacy get_current_balance
                # swallows IronPython attribute errors and returns 0.0,
                # which would synthesize a bogus starter row.
                actual_due = None
                try:
                    bsql = ("SELECT SUM(ISNULL(IndDue, 0)) AS B "
                            "FROM dbo.TransactionSummary "
                            "WHERE PeopleId = {0} AND OrganizationId = {1}"
                            .format(int(people_id), int(org_id)))
                    bres = list(q.QuerySql(bsql))
                    if bres:
                        raw_b = getattr(bres[0], 'B', None)
                        if raw_b is not None:
                            actual_due = float(raw_b)
                except:
                    actual_due = None
                # Only synthesize the starter row when:
                #   * lookup succeeded (not None) AND
                #   * there's a real gap (> 1 cent), AND
                #   * the gap goes the right direction -- a positive
                #     gap (actual > running) means the visible rows
                #     UNDER-count what they owe (typical: implicit reg
                #     fee). A negative gap would mean the visible rows
                #     OVER-count, which usually means our lookup is the
                #     unreliable one (returned 0 spuriously), so we
                #     skip rather than mislead.
                if actual_due is not None and (actual_due - running) > 0.01:
                    gap = actual_due - running
                    starter = {
                        'tranDate':       'Starting balance',
                        'transactionId':  '',
                        'kind':           'Prior',
                        'description':    'Registration fee or prior balance not in transaction history',
                        'note':           'Prior balance',
                        'message':        '',
                        'amount':         abs(gap),
                        'balanceDelta':   gap,
                        'isPayment':      gap < 0,
                        'runningBalance': gap,
                    }
                    for row in rows:
                        row['runningBalance'] += gap
                    rows.insert(0, starter)
                    running = actual_due
                return self.create_json_response(True, "ok", {
                    'charges': rows,
                    'endingBalance': running,
                })
            except Exception as e:
                return self.create_json_response(False, "charge_details error: " + str(e))

        def get_transaction_history(self, people_id):
            """Get transaction history for a person"""
            sql = '''
            SELECT 
                p.Name Person, pro.Name as Program, d.Name as Division, o.OrganizationName OrgName,
                FORMAT(ts.TranDate, 'yyyy-MM-dd') as TranDate, ts.TotDue as TSAmount,
                FORMAT(t.TransactionDate, 'yyyy-MM-dd') as TransactionDate, t.amtdue as Amount,
                CASE WHEN ts.TotDue > 0 THEN 'Church Received'
                     WHEN ts.TotDue < 0 THEN 'Person Received'
                     ELSE 'Zero Amount' END TransactionDirection,
                t.TransactionId, t.Message, t.Description,
                CASE WHEN [message] like 'CHK%' THEN 'Check'
                     WHEN [message] like 'CSH%' THEN 'Cash'
                     WHEN [message] like 'Response%' THEN 'Credit Card'
                     WHEN [message] like 'FEE%' THEN 'Church Adjustment'
                     WHEN [transactionid] like 'Coupon%' THEN 'Coupon'
                     ELSE 'Unknown' END TransactionType
            FROM TransactionSummary ts
            LEFT JOIN [Transaction] t on ts.RegId = t.OriginalId
            INNER JOIN [People] p on ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE ts.PeopleId = {0} AND t.amtdue <> 0
            ORDER BY t.TransactionDate DESC
            '''.format(people_id)
            return q.QuerySql(sql)

        # Payment processing methods
        def process_payment_link(self):
            """Process payment link request.

            channel = 'email' (default), 'text', or 'both'. Email uses
            the configured sender + template; text uses model.SendSms
            with the first SmsGroup row (matching the FMC Member
            Manager pattern). Returns a single success/fail with a
            channel-aware message.
            """
            try:
                payer_id = str(getattr(model.Data, 'pid', ''))
                org_id = str(getattr(model.Data, 'PaymentOrg', ''))
                amount = str(getattr(model.Data, 'PayFee', ''))
                payer_name = str(getattr(model.Data, 'payerName', ''))
                cc_emails = str(getattr(model.Data, 'cc_emails', ''))
                channel = str(getattr(model.Data, 'channel', 'email') or 'email').strip().lower()
                if channel not in ('email', 'text', 'both'):
                    channel = 'email'

                if not payer_id or not org_id or not amount:
                    return self.create_json_response(False, "Missing required payment information")

                # Get email sender details
                program_id = str(getattr(model.Data, 'ProgramID', self.program_id))
                email_details = self.get_email_details(program_id)

                if not email_details:
                    return self.create_json_response(False, "Email configuration not found for program")

                # Get current balance
                previous_due = self.get_current_balance(payer_id, org_id)

                # Generate payment link
                paylink = model.GetPayLink(int(payer_id), int(org_id))
                paylinkauth = model.GetAuthenticatedUrl(int(org_id), paylink, True)

                if str(paylinkauth).split('/')[-1].lower() == 'none':
                    return self.create_json_response(False, "Unable to generate payment link")

                # Calculate total due
                total_due = float(previous_due) + float(amount)

                sent = []
                errors = []

                # ----- TEXT -----
                if channel in ('text', 'both'):
                    try:
                        sms_group = q.QuerySqlInt("SELECT TOP 1 ID FROM SmsGroups")
                        if not sms_group:
                            errors.append("No SmsGroup configured")
                        else:
                            sms_title = (email_details.get('email_title') or 'Payment Link')
                            sms_body = ("Hi " + str(payer_name) + ", please pay your "
                                        "$" + '{:,.2f}'.format(float(total_due)) + " balance here: "
                                        + str(paylinkauth) + "  (This is a one-way message; we will not receive replies.)")
                            model.SendSms('PeopleId = ' + str(int(payer_id)), int(sms_group), sms_title, sms_body)
                            sent.append('text')
                    except Exception as se:
                        errors.append("Text failed: " + str(se))

                # ----- EMAIL -----
                if channel in ('email', 'both'):
                    try:
                        message = self.build_payment_email(
                            payer_name, amount, previous_due, total_due,
                            paylinkauth, email_details
                        )
                        model.Email(
                            int(payer_id),
                            int(email_details.get('sender_id') or 0),
                            email_details.get('sender_email', ''),
                            email_details.get('sender_alias', 'Payment System'),
                            email_details.get('email_title', 'Payment Notification'),
                            message,
                            cc_emails
                        )
                        sent.append('email')
                    except Exception as ee:
                        errors.append("Email failed: " + str(ee))

                if sent and not errors:
                    return self.create_json_response(True, "Payment link sent via " + ' + '.join(sent))
                if sent and errors:
                    return self.create_json_response(True, "Sent via " + ' + '.join(sent) + " (but: " + '; '.join(errors) + ")")
                if errors:
                    return self.create_json_response(False, '; '.join(errors))
                return self.create_json_response(False, "Nothing sent")
            except Exception as e:
                return self.create_json_response(False, "Error processing payment link: " + str(e))

        def process_zero_out_refund(self):
            """One-click clear of a refund-driven outstanding balance.

            For SINGLE-person registrations: applies an AdjustFee
            credit equal to the person's IndDue.

            For MULTI-person registrations (e.g. one payer covered the
            whole family, then everyone was refunded together): clears
            EVERY co-registrant on the same RegId. Without this, a
            family of four would need four separate Zero Out clicks
            and the IsRefundDriven flag would stop firing after the
            first one (the RegId total IndDue would no longer match
            the refund). Single transaction, single audit-line per
            person ('ADJ|Refund cleared: not actually due').

            Original payment + refund rows are preserved in the
            Transaction table -- only clean ADJ| audit rows are added.
            """
            try:
                payer_id = str(getattr(model.Data, 'pid', '') or '').strip()
                org_id   = str(getattr(model.Data, 'PaymentOrg', '') or '').strip()
                amount   = str(getattr(model.Data, 'amount', '') or '').strip()
                if not payer_id or not org_id:
                    return self.create_json_response(False, "Missing payer or involvement")
                try:
                    amt = float(amount)
                except:
                    return self.create_json_response(False, "Invalid amount")
                if amt <= 0:
                    return self.create_json_response(False, "Amount must be greater than zero")
                # Cross-check the clicked person's IndDue first.
                try:
                    bres = list(q.QuerySql(
                        "SELECT SUM(ISNULL(IndDue, 0)) AS B FROM dbo.TransactionSummary "
                        "WHERE PeopleId = {0} AND OrganizationId = {1}"
                        .format(int(payer_id), int(org_id))
                    ))
                    actual_due = float(getattr(bres[0], 'B', 0) or 0) if bres else 0.0
                except:
                    actual_due = None
                if actual_due is None or abs(actual_due - amt) > 0.01:
                    return self.create_json_response(False,
                        "Outstanding has changed since this page loaded "
                        "(now $" + ('{:,.2f}'.format(actual_due) if actual_due is not None else '?') +
                        " vs $" + '{:,.2f}'.format(amt) + " on screen). Refresh and try again.")
                # Find ALL co-registrants on the same RegId(s) with
                # IndDue > 0. A multi-person family registration will
                # produce >1 row here; a solo registration just 1.
                try:
                    coreg_sql = """
                        SELECT ts.PeopleId, ts.IndDue, ISNULL(p.Name, '') AS PersonName
                        FROM TransactionSummary ts
                        INNER JOIN People p ON p.PeopleId = ts.PeopleId
                        WHERE ts.RegId IN (
                            SELECT RegId FROM TransactionSummary
                            WHERE PeopleId = {0} AND OrganizationId = {1}
                              AND IsLatestTransaction = 1
                        )
                          AND ts.IsLatestTransaction = 1
                          AND ts.IndDue > 0
                    """.format(int(payer_id), int(org_id))
                    coregs = list(q.QuerySql(coreg_sql))
                except Exception as ce:
                    return self.create_json_response(False, "Failed to look up co-registrants: " + str(ce))
                if not coregs:
                    return self.create_json_response(False, "No outstanding rows found for this registration.")
                cleared = []
                errors = []
                for r in coregs:
                    try:
                        cpid = int(getattr(r, 'PeopleId', 0) or 0)
                        camt = float(getattr(r, 'IndDue', 0) or 0)
                        cname = str(getattr(r, 'PersonName', '') or '')
                        if cpid and camt > 0:
                            model.AdjustFee(cpid, int(org_id), camt,
                                            'ADJ|Refund cleared: not actually due (auto-zeroed)')
                            cleared.append(cname + ' $' + '{:,.2f}'.format(camt))
                    except Exception as ae:
                        errors.append(str(getattr(r, 'PersonName', '?')) + ': ' + str(ae))
                if not cleared:
                    return self.create_json_response(False,
                        "Nothing cleared. " + ('; '.join(errors) if errors else ''))
                msg = "Cleared " + str(len(cleared)) + " row(s)"
                if len(cleared) > 1:
                    msg += " on this registration: " + ', '.join(cleared)
                else:
                    msg += " (" + cleared[0] + ")"
                if errors:
                    msg += ". WARNING: " + '; '.join(errors)
                msg += ". Refund audit preserved in Payment History."
                return self.create_json_response(True, msg)
            except Exception as e:
                return self.create_json_response(False, "Error zeroing balance: " + str(e))

        def process_adjust_balance(self):
            """Adjust a payer's balance for a given involvement.

            direction = 'charge' (increase balance / add fee)
                        | 'credit' (decrease balance / apply credit)
            amount    = positive decimal entered by the staffer.
            note      = free-text reason; saved as 'ADJ|<note>' so the
                        inline breakdown shows it as an Adjustment.
            """
            try:
                payer_id  = str(getattr(model.Data, 'pid', '') or '').strip()
                org_id    = str(getattr(model.Data, 'PaymentOrg', '') or '').strip()
                direction = str(getattr(model.Data, 'direction', 'charge') or 'charge').strip().lower()
                amount    = str(getattr(model.Data, 'amount', '') or '').strip()
                note      = str(getattr(model.Data, 'note', '') or '').strip()
                if not payer_id or not org_id:
                    return self.create_json_response(False, "Missing payer or involvement")
                if direction not in ('charge', 'credit'):
                    return self.create_json_response(False, "Direction must be 'charge' or 'credit'")
                try:
                    amt = float(amount)
                    if amt <= 0:
                        return self.create_json_response(False, "Amount must be greater than zero")
                except:
                    return self.create_json_response(False, "Invalid amount")
                if not note:
                    return self.create_json_response(False, "Please describe the reason for this adjustment")
                # TouchPoint convention for AdjustFee (matches the 'amt'
                # column on Transaction):
                #   + amount = payment-shaped, reduces balance (credit)
                #   - amount = charge-shaped, increases balance
                # We verified empirically that the inverse of my first
                # guess was correct.
                signed = -amt if direction == 'charge' else amt
                # Prefix the description with ADJ| so the inline charge
                # breakdown labels it as an Adjustment.
                description = 'ADJ|' + note
                try:
                    model.AdjustFee(int(payer_id), int(org_id), signed, description)
                except Exception as ae:
                    return self.create_json_response(False, "AdjustFee failed: " + str(ae))
                # New balance for the response toast.
                try:
                    new_balance = float(self.get_current_balance(payer_id, org_id) or 0)
                except:
                    new_balance = None
                verb = 'Charged' if direction == 'charge' else 'Credited'
                msg  = verb + ' ${:,.2f}'.format(amt)
                if new_balance is not None:
                    msg += '. New balance: ${:,.2f}'.format(new_balance)

                # ----- Optional email notification -----
                send_email = str(getattr(model.Data, 'send_email', 'false') or 'false').lower() in ('true', '1', 'yes', 'on')
                email_status = ''
                if send_email:
                    try:
                        person = model.GetPerson(int(payer_id))
                        person_email = str(getattr(person, 'EmailAddress', '') or '') if person else ''
                        person_name = ''
                        if person:
                            try:    person_name = str(person.Name or '')
                            except: person_name = ''
                        if not person_email:
                            email_status = " (no email on file -- not sent)"
                        else:
                            settings = pm_load_settings()
                            program_id = str(getattr(model.Data, 'ProgramID', self.program_id) or '')
                            email_details = self.get_email_details(program_id) or DEFAULT_EMAIL_SENDER
                            # Build a self-contained notification body
                            # (not template-driven -- this is a new flow
                            # distinct from receipts / pay-link emails).
                            paylink_html = ''
                            if new_balance and new_balance > 0:
                                try:
                                    pl = model.GetPayLink(int(payer_id), int(org_id))
                                    pla = model.GetAuthenticatedUrl(int(org_id), pl, True)
                                    if str(pla).split('/')[-1].lower() != 'none':
                                        paylink_html = ('<p style="text-align:center;margin:20px 0;">'
                                                        '<a href="' + str(pla) + '" style="display:inline-block;'
                                                        'background:#1f4e79;color:#fff;text-decoration:none;'
                                                        'padding:10px 22px;border-radius:6px;font-weight:600;">Pay Now &raquo;</a></p>')
                                except:
                                    pass
                            badge_bg = '#fff4ce' if direction == 'charge' else '#d4edda'
                            badge_fg = '#7a5c00' if direction == 'charge' else '#155724'
                            badge_lbl = ('Charge added' if direction == 'charge' else 'Credit applied')
                            bal_color = '#7a5c00' if (new_balance is None or new_balance > 0) else '#155724'
                            bal_str = '${:,.2f}'.format(float(new_balance) if new_balance is not None else 0)
                            body_html = (
                                '<div style="font-family:Segoe UI,Tahoma,sans-serif;color:#333;max-width:600px;">'
                                '<h2 style="color:#1f4e79;margin:0 0 16px;">Account Adjustment</h2>'
                                '<p>Hi ' + (person_name or 'there') + ',</p>'
                                '<p>An adjustment was made to your account. Details:</p>'
                                '<table style="border-collapse:collapse;margin:12px 0;">'
                                '<tr><td style="padding:6px 14px 6px 0;color:#666;">Type:</td>'
                                '<td><span style="background:' + badge_bg + ';color:' + badge_fg + ';padding:2px 10px;border-radius:10px;font-size:12px;font-weight:700;">' + badge_lbl + '</span></td></tr>'
                                '<tr><td style="padding:6px 14px 6px 0;color:#666;">Amount:</td><td><strong>${:,.2f}</strong></td></tr>'.format(amt) +
                                '<tr><td style="padding:6px 14px 6px 0;color:#666;">Reason:</td><td>' + note + '</td></tr>'
                                '<tr><td style="padding:6px 14px 6px 0;color:#666;">Current balance:</td><td style="color:' + bal_color + ';font-weight:700;">' + bal_str + '</td></tr>'
                                '</table>'
                                + paylink_html +
                                '<p style="font-size:12px;color:#666;margin-top:24px;">' +
                                str(email_details.get('sender_alias', 'Payment System')) + '<br>' +
                                str(email_details.get('sender_email', '')) + '<br>' +
                                str(email_details.get('sender_phone', '')) +
                                '</p></div>'
                            )
                            subject = (str(email_details.get('email_title') or 'Account Adjustment')
                                       + ' -- ' + badge_lbl)
                            try:
                                model.Email(
                                    int(payer_id),
                                    int(email_details.get('sender_id') or 0),
                                    email_details.get('sender_email', ''),
                                    email_details.get('sender_alias', 'Payment System'),
                                    subject,
                                    body_html,
                                    ''  # no cc
                                )
                                email_status = ' (emailed ' + person_email + ')'
                            except Exception as ee:
                                email_status = ' (adjust applied but email failed: ' + str(ee) + ')'
                    except Exception as eo:
                        email_status = ' (adjust applied but email failed: ' + str(eo) + ')'

                return self.create_json_response(True, msg + email_status, {'newBalance': new_balance})
            except Exception as e:
                return self.create_json_response(False, "Error adjusting balance: " + str(e))

        def process_payment_record(self):
            """
            Process payment recording
            
            Uses template: PAYMENT_CONFIRMATION_TEMPLATE_NAME
            Variables: name, chargeNotes, previousDue, newTotalDue, sender_alias, sender_phone, sender_email
            """
            try:
                # Extract form data with safe fallbacks
                payer_id = str(getattr(model.Data, 'pid', ''))
                org_id = str(getattr(model.Data, 'PaymentOrg', ''))
                payment_amount = str(getattr(model.Data, 'PaidAmount', ''))
                payment_type = str(getattr(model.Data, 'PaymentType', CHECK_PAYMENT_TYPE))
                payment_desc = str(getattr(model.Data, 'PaymentDescription', ''))
                payer_name = str(getattr(model.Data, 'payerName', ''))
                cc_emails = str(getattr(model.Data, 'cc_emails', ''))
                
                # Validate required fields
                if not all([payer_id, org_id, payment_amount, payment_type, payment_desc]):
                    return self.create_json_response(False, "Missing required payment information")
                
                # Validate amount format
                if payment_amount != "" and str(payment_amount).replace(".", "").isdigit():
                    # Payment comes in as a positive number but needs to be negative in the database
                    amt = -float(payment_amount)
                    PayType = 'Payment'
                    payment_amount_display = '${0:.2f}'.format(abs(float(amt))).replace("$-", "-$")
                    chargeNotes = payment_type + ' ' + payment_amount_display + ': ' + payment_desc
                else:
                    return self.create_json_response(False, "Invalid payment amount")
                
                # Get email configuration
                program_id = str(getattr(model.Data, 'ProgramID', self.program_id))
                email_details = self.get_email_details(program_id)
                
                # Prepare transaction description
                messageDescription = payment_type + payment_desc
                
                # Get previous balance
                previousDue = "<br />" + '$0.00' + "........Previous Credit Amount"
                IndDue = 0
                
                # Query current balance
                try:
                    inddue_list = q.QuerySql(
                        "SELECT Sum(IndDue) as IndDue, Sum(TotDue) as TotDue " +
                        "FROM dbo.TransactionSummary " +
                        "WHERE PeopleId = {0} AND OrganizationId = {1}".format(payer_id, org_id)
                    )
                    
                    if inddue_list:
                        for pc in inddue_list:
                            # Handle different balance scenarios
                            if pc.IndDue is None or pc.IndDue == 0:
                                IndDue = float(0)
                            elif pc.IndDue > 0 or pc.IndDue < 0:
                                IndDue = float(pc.IndDue)
                            
                            # Create initial transaction if needed
                            if pc.TotDue is None:
                                model.AddTransaction(int(payer_id), int(org_id), 0, "Initial charge of $0")
                            
                            # Format previous balance for display
                            previousDue = "<br />" + '${:,.2f}'.format(IndDue + 0.00).replace("$-","-$") + "........Previous Balance"
                except Exception as query_error:
                    print("<!-- Error querying balance: " + str(query_error) + " -->")
                    # Continue with default values
                
                # Check for overpayment
                if IndDue + float(amt) < 0:
                    PayType = 'Payment'
                
                # Process payment
                if PayType == 'Payment':
                    # Calculate new total due
                    newTotalDue = float(IndDue) + amt
                    newTotalDue_display = '${:,.2f}'.format(float(newTotalDue)*1.0 + 0.00).replace("$-","-$")
                    
                    # Add transaction to database
                    try:
                        # In DB: positive = charge, negative = payment
                        transamount = -amt  # Negate the already negative amount to make it positive
                        transaction_id = model.AddTransaction(int(payer_id), int(org_id), transamount, messageDescription)
                        
                        if not transaction_id:
                            return self.create_json_response(False, "Transaction not recorded - please check credentials")
                    except Exception as e:
                        return self.create_json_response(False, "Failed to record transaction: " + str(e))
                    
                    # Send confirmation email
                    try:
                        # Try to load configured template
                        try:
                            message = model.HtmlContent(PAYMENT_CONFIRMATION_TEMPLATE_NAME)
                        except Exception as template_error:
                            # Fall back to default template if configured one not found
                            print("<!-- Template not found: " + str(template_error) + " -->")
                            message = DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE
                        
                        # Format template with payment details
                        formatted_message = message.format(
                            name=payer_name, 
                            chargeNotes=chargeNotes, 
                            previousDue=previousDue,
                            newTotalDue=newTotalDue_display, 
                            sender_alias=email_details.get('sender_alias', ''), 
                            sender_phone=email_details.get('sender_phone', ''), 
                            sender_email=email_details.get('sender_email', '')
                        )
                        
                        # Send the email
                        try:
                            model.Email(
                                int(payer_id), 
                                int(email_details.get('sender_id', 3134)), 
                                email_details.get('sender_email', 'noreply@church.com'),
                                email_details.get('sender_alias', 'Payment System'), 
                                EMAIL_SUBJECT_PREFIX + "Payment Received", 
                                formatted_message, 
                                cc_emails
                            )
                        except Exception as email_error:
                            # Log but don't fail the transaction if email fails
                            print("<!-- Email error: " + str(email_error) + " -->")
                    except Exception as template_error:
                        # Log but don't fail if template processing fails
                        print("<!-- Template error: " + str(template_error) + " -->")
                
                # Return success response
                return self.create_json_response(True, "Payment recorded successfully")
                
            except Exception as e:
                return self.create_json_response(False, "Error recording payment: " + str(e))

        def process_resend_email(self):
            """Process email resend request"""
            try:
                message_id = str(getattr(model.Data, 'messageId', ''))
                people_id = str(getattr(model.Data, 'PeopleId', ''))
                
                if not message_id or not people_id:
                    return self.create_json_response(False, "Missing message ID or people ID")
                
                # Use the same SQL as your original PM_EmailPreview
                sql = '''
                Select eq.Subject
                    ,eq.Body
                    ,eq.Sent
                    ,eq.FromName
                    ,eq.Id AS [messageId]
                    ,eqt.PeopleId
                    ,eq.QueuedBy
                    ,eq.FromAddr
                from EmailQueue eq
                left join dbo.EmailQueueTo eqt on (eqt.id = eq.id and eqt.PeopleId = {1})
                where eq.Id = {0}
                '''.format(message_id, people_id)
                
                email_data = q.QuerySql(sql)
                if not email_data or len(email_data) == 0:
                    return self.create_json_response(False, "Original email not found")
                
                # Use a for loop like your original code
                for a in email_data:
                    email_title = 'Email Copy - ' + str(self.safe_get_attr(a, 'Subject', 'No Subject'))
                    email_body = '<H2>Email Copy - Originally Sent on <i>' + str(self.safe_get_attr(a, 'Sent', '')) + '</i></H2><br />' + str(self.safe_get_attr(a, 'Body', 'No content'))
                    
                    try:
                        model.Email(
                            int(people_id), 
                            self.safe_get_attr(a, 'QueuedBy', 3134), 
                            str(self.safe_get_attr(a, 'FromAddr', 'noreply@church.com')),
                            str(self.safe_get_attr(a, 'FromName', 'Payment System')), 
                            email_title, 
                            email_body
                        )
                        return self.create_json_response(True, "Email copy sent successfully")
                    except Exception as e:
                        return self.create_json_response(False, "Failed to send email: " + str(e))
                        
            except Exception as e:
                return self.create_json_response(False, "Error processing email resend: " + str(e))

        def get_email_details(self, program_id):
            """Get email configuration for program. Routes through the new
            Settings store; falls back to module DEFAULT_EMAIL_SENDER if
            the Settings UI hasn't been touched yet."""
            try:
                settings = pm_load_settings()
                senders = settings.get('senders', {}) or {}
                pid_key = str(program_id) if program_id is not None else ''
                if pid_key in senders and isinstance(senders[pid_key], dict):
                    return senders[pid_key]
                return settings.get('defaultSender') or DEFAULT_EMAIL_SENDER
            except:
                return DEFAULT_EMAIL_SENDER

        # ==============================================================
        # RECEIPTS REPRINT -- find past check/cash payments, re-fire email
        # ==============================================================
        def process_find_receipts(self):
            """Date-range search for check/cash payments. Returns JSON
            list of {transactionId, person, peopleId, email, orgName,
            amount, transactionDate, paymentType, source, description}.

            Source detection:
              * 'pm'       -- Message looks like 'CHK|...' or 'CSH|...'.
                              Payment Manager writes the pipe; only these
                              are safe to reprint a receipt for, since
                              they're the only payments that had an
                              original receipt fire from this tool.
              * 'external' -- anything else (TouchPoint batch entry,
                              other apps). Receipts are NOT shown for
                              these unless include_external=true.
            """
            try:
                settings = pm_load_settings()
                pm_chk = (settings.get('checkPaymentType') or CHECK_PAYMENT_TYPE).replace("'", "''")
                pm_csh = (settings.get('cashPaymentType')  or CASH_PAYMENT_TYPE).replace("'", "''")
                date_from = str(getattr(model.Data, 'date_from', '')).strip()
                date_to   = str(getattr(model.Data, 'date_to', '')).strip()
                # rcpt_types is a comma-separated multi-select. Legacy
                # rcpt_type ('check'|'cash'|'both') is honored for any
                # callers that haven't upgraded.
                raw_types = str(getattr(model.Data, 'rcpt_types', '')).strip().lower()
                if not raw_types:
                    legacy = str(getattr(model.Data, 'rcpt_type', 'both')).strip().lower()
                    if legacy == 'both':
                        raw_types = 'check,cash'
                    else:
                        raw_types = legacy
                selected_types = set([t.strip() for t in raw_types.split(',') if t.strip()])
                include_external = str(getattr(model.Data, 'include_external', '')).lower() in ('true', '1', 'yes', 'on')
                # Optional substring filters. Either can be empty.
                payer_filter = str(getattr(model.Data, 'payer', '') or '').strip()
                org_filter   = str(getattr(model.Data, 'involvement', '') or '').strip()
                # Last-4 search: matches Transaction.LastFourCC or
                # LastFourACH exactly (after stripping non-digits).
                last4_filter = ''.join(c for c in (str(getattr(model.Data, 'last4', '') or '')) if c.isdigit())
                # Limit to last 4 digits if more typed
                if len(last4_filter) > 4:
                    last4_filter = last4_filter[-4:]
                if not date_from or not date_to:
                    return self.create_json_response(False, "Both From and To dates are required")
                # Build the type filter.
                #
                # "Refund" is a sixth checkbox that toggles whether
                # negative-amt rows (money-out reversals) are included.
                # When unchecked, every other payment-type clause adds
                # 'AND amt > 0' so reversals are kept out -- pure
                # payment receipts only. When checked alongside another
                # type, both payments AND reversals of that type show.
                # When checked alone, refunds of any type show.
                want_refund = 'refund' in selected_types
                # Sign filter applied to the non-refund clauses. When
                # refund is unchecked we restrict to positive amt; when
                # refund is checked we leave the sign open so each
                # type's checkbox effectively means "any direction".
                amt_sign = '' if want_refund else ' AND t.amt > 0'
                type_clauses = []
                if 'check' in selected_types:
                    if include_external:
                        type_clauses.append("(t.[Message] LIKE 'CHK%'" + amt_sign + ")")
                    else:
                        type_clauses.append("(t.[Message] LIKE '" + pm_chk + "%'" + amt_sign + ")")
                if 'cash' in selected_types:
                    if include_external:
                        type_clauses.append("(t.[Message] LIKE 'CSH%'" + amt_sign + ")")
                    else:
                        type_clauses.append("(t.[Message] LIKE '" + pm_csh + "%'" + amt_sign + ")")
                if 'credit' in selected_types or 'creditcard' in selected_types or 'cc' in selected_types:
                    type_clauses.append("(t.[Message] LIKE 'Response%'" + amt_sign + ")")
                if 'coupon' in selected_types:
                    # Coupons identified by TransactionId; sign filter
                    # applies the same way.
                    type_clauses.append("(t.TransactionId LIKE 'Coupon%'" + amt_sign + ")")
                if 'other' in selected_types:
                    # Money-in entries that don't carry one of the
                    # standard prefixes. Always amt > 0 here regardless
                    # of refund checkbox -- refunds of unprefixed
                    # entries are caught by the refund clause below
                    # via the payment-shape catch-all.
                    type_clauses.append(
                        "(t.amt > 0 "
                        "  AND t.[Message] NOT LIKE 'CHK%' "
                        "  AND t.[Message] NOT LIKE 'CSH%' "
                        "  AND t.[Message] NOT LIKE 'Response%' "
                        "  AND t.TransactionId NOT LIKE 'Coupon%' "
                        "  AND ISNULL(t.[Message], '') NOT LIKE 'ADJ|%' "
                        "  AND ISNULL(t.[Message], '') NOT LIKE 'FEE|%' "
                        "  AND ISNULL(t.[Message], '') NOT LIKE 'variable%')"
                    )
                if want_refund:
                    # Refund = negative amt on a payment-shaped row.
                    # When the user picks ONLY refund, this is the only
                    # clause that fires; when they pick refund alongside
                    # CC (etc.), the CC clause already has no amt sign
                    # filter so refunds for that method are included --
                    # but this clause also adds refunds for any payment
                    # methods the user didn't explicitly check. Avoid
                    # double-counting by skipping methods already
                    # selected.
                    refund_methods = []
                    if 'check' not in selected_types:
                        refund_methods.append("t.[Message] LIKE 'CHK%'")
                    if 'cash' not in selected_types:
                        refund_methods.append("t.[Message] LIKE 'CSH%'")
                    if not any(k in selected_types for k in ('credit', 'creditcard', 'cc')):
                        refund_methods.append("t.[Message] LIKE 'Response%'")
                    if 'coupon' not in selected_types:
                        refund_methods.append("t.TransactionId LIKE 'Coupon%'")
                    if refund_methods:
                        type_clauses.append(
                            "(t.amt < 0 AND (" + " OR ".join(refund_methods) + "))"
                        )
                if not type_clauses:
                    return self.create_json_response(False, "Pick at least one payment type")
                type_sql = '(' + ' OR '.join(type_clauses) + ')'
                # Sanitize dates -- only ISO yyyy-mm-dd accepted.
                def _iso(s):
                    return bool(re.match(r'^\d{4}-\d{2}-\d{2}$', s or ''))
                if not (_iso(date_from) and _iso(date_to)):
                    return self.create_json_response(False, "Dates must be in YYYY-MM-DD format")
                # Use t.amt (the actual payment amount) instead of
                # t.amtdue -- amtdue is NULL or 0 for many rows
                # (especially CC and Coupon), which would silently drop
                # them. MinistryDepositReport uses the same convention.
                sql = """
                    SELECT
                        t.TransactionId,
                        ISNULL(NULLIF(t.Name, ''), ISNULL(p.Name, '')) AS PersonName,
                        ISNULL(p.PeopleId, 0)     AS PeopleId,
                        ISNULL(p.EmailAddress, '') AS Email,
                        ISNULL(o.OrganizationName, '') AS OrgName,
                        ISNULL(o.OrganizationId, 0) AS OrgId,
                        t.amt                     AS Amount,
                        FORMAT(t.TransactionDate, 'yyyy-MM-dd HH:mm') AS TranDate,
                        FORMAT(t.TransactionDate, 'yyyy-MM-dd') AS TranDateOnly,
                        CASE WHEN t.[Message] LIKE 'CHK%' AND t.amt < 0 THEN 'Check Refund'
                             WHEN t.[Message] LIKE 'CSH%' AND t.amt < 0 THEN 'Cash Refund'
                             WHEN t.[Message] LIKE 'Response%' AND t.amt < 0 THEN 'Refund'
                             WHEN t.[Message] LIKE 'CHK%' THEN 'Check'
                             WHEN t.[Message] LIKE 'CSH%' THEN 'Cash'
                             WHEN t.[Message] LIKE 'Response%' THEN 'Credit Card'
                             WHEN t.TransactionId LIKE 'Coupon%' THEN 'Coupon'
                             WHEN t.amt > 0 THEN 'Other Payment'
                             ELSE 'Other' END AS PaymentType,
                        ISNULL(t.[Message], '')   AS Message,
                        ISNULL(t.[Description], '') AS Description,
                        ISNULL(t.LastFourCC, '')  AS LastFourCC,
                        ISNULL(t.LastFourACH, '') AS LastFourACH,
                        -- All people linked to this registration. Lets
                        -- the UI explain why several rows share a date /
                        -- amount / payer -- a family registered multiple
                        -- kids on one transaction.
                        ISNULL(STUFF((
                            SELECT ', ' + p2.Name
                            FROM TransactionSummary ts2
                            JOIN People p2 ON p2.PeopleId = ts2.PeopleId
                            WHERE ts2.RegId = t.OriginalId
                              AND ts2.IsLatestTransaction = 1
                            FOR XML PATH('')
                        ), 1, 2, ''), '') AS RegistrantNames
                    FROM [Transaction] t
                    -- OUTER APPLY (TOP 1) prevents row-multiplication when
                    -- a transaction is linked to several TransactionSummary
                    -- rows -- e.g. multi-person registrations, multi-use
                    -- coupons, recurring adjustments. Locally
                    -- 'Adjustment (12431)' multiplied to 12 rows under a
                    -- plain LEFT JOIN.
                    OUTER APPLY (
                        SELECT TOP 1 ts.PeopleId, ts.OrganizationId
                        FROM TransactionSummary ts
                        WHERE ts.RegId = t.OriginalId
                          AND ts.IsLatestTransaction = 1
                        ORDER BY ts.PeopleId
                    ) ts
                    LEFT JOIN People p ON p.PeopleId = ts.PeopleId
                    LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
                    WHERE {0}
                      AND t.TransactionDate >= '{1}'
                      AND t.TransactionDate <  DATEADD(day, 1, '{2}')
                      AND t.amt <> 0
                      AND t.TransactionId IS NOT NULL
                      AND t.voided IS NULL
                      {3}
                      {4}
                    ORDER BY t.TransactionDate DESC
                """.format(type_sql, date_from, date_to,
                           ("AND (p.Name LIKE '%" + payer_filter.replace("'", "''") + "%' OR t.Name LIKE '%" + payer_filter.replace("'", "''") + "%')") if payer_filter else "",
                           ("AND o.OrganizationName LIKE '%" + org_filter.replace("'", "''") + "%'") if org_filter else "")
                # Last-4 narrows on Transaction.LastFourCC or LastFourACH
                # AFTER the OUTER APPLY structure -- inject before the
                # ORDER BY. Doing it as a string splice keeps the
                # OUTER APPLY block readable.
                if last4_filter:
                    sql = sql.replace(
                        "ORDER BY t.TransactionDate DESC",
                        "AND (t.LastFourCC = '" + last4_filter + "' OR t.LastFourACH = '" + last4_filter + "')\n                    ORDER BY t.TransactionDate DESC"
                    )
                rows = []
                for r in q.QuerySql(sql):
                    msg = str(self.safe_get_attr(r, 'Message', '') or '')
                    txid = str(self.safe_get_attr(r, 'TransactionId', '') or '')
                    # Source = 'pm' when message begins with PM's exact
                    # signature (CHK|... / CSH|...). Anything else came
                    # from elsewhere -- batch entry, sync, etc.
                    src = 'pm' if (msg.startswith(pm_chk) or msg.startswith(pm_csh)) else 'external'
                    # Extract the staff note from the Message so the row
                    # list can show "Added banquet via phone call" instead
                    # of a hex TransactionId. Skips CC gateway noise.
                    note = ''
                    if '|' in msg:
                        note = msg.split('|', 1)[1].strip()
                    elif msg and not msg.startswith('Response'):
                        note = msg.strip()
                    rows.append({
                        'transactionId': txid,
                        'person':        self.safe_get_attr(r, 'PersonName', ''),
                        'peopleId':      self.safe_get_attr(r, 'PeopleId', 0),
                        'email':         self.safe_get_attr(r, 'Email', ''),
                        'orgName':       self.safe_get_attr(r, 'OrgName', ''),
                        'orgId':         self.safe_get_attr(r, 'OrgId', 0),
                        'amount':        float(self.safe_get_attr(r, 'Amount', 0) or 0),
                        'tranDate':      self.safe_get_attr(r, 'TranDate', ''),
                        'tranDateOnly':  self.safe_get_attr(r, 'TranDateOnly', ''),
                        'paymentType':   self.safe_get_attr(r, 'PaymentType', ''),
                        'registrants':   self.safe_get_attr(r, 'RegistrantNames', ''),
                        'message':       msg,
                        'note':          note,
                        'source':        src,
                        'description':   self.safe_get_attr(r, 'Description', ''),
                        'last4cc':       self.safe_get_attr(r, 'LastFourCC', ''),
                        'last4ach':      self.safe_get_attr(r, 'LastFourACH', ''),
                    })
                # Diagnostic counts so the empty-state can explain why
                # nothing showed up. We count UNFILTERED (no amtdue gate
                # or type gate) so the staffer can tell the difference
                # between "no transactions in this range" and "lots of
                # transactions but none with a payment amount." We also
                # surface the most recent transaction date so it's
                # obvious when a test/QA database is just stale.
                # Capture raw model.Data values verbatim so the diagnostic
                # can show exactly what arrived from the browser (helps
                # catch cases where a value silently becomes empty).
                try:
                    raw_df = repr(getattr(model.Data, 'date_from', None))
                except: raw_df = '<read-error>'
                try:
                    raw_dt = repr(getattr(model.Data, 'date_to', None))
                except: raw_dt = '<read-error>'
                diag = {'totalTxns': 0, 'chkCount': 0, 'cshCount': 0, 'ccCount': 0,
                        'couponCount': 0, 'otherCount': 0, 'pmCount': 0, 'extCount': 0,
                        'nonZeroAmt': 0, 'latestTxnDate': '',
                        'tableTotal': 0, 'tableProbeErr': '',
                        'dateFromEcho': date_from, 'dateToEcho': date_to,
                        'rawDateFrom': raw_df, 'rawDateTo': raw_dt}
                if not rows:
                    # ALWAYS run the table probes first -- if these fail,
                    # we want to KNOW (regardless of whether the filtered
                    # diagnostic succeeds). Each in its own try.
                    try:
                        tc = list(q.QuerySql("SELECT COUNT(*) AS TableTotal FROM [Transaction]"))
                        if tc:
                            diag['tableTotal'] = int(self.safe_get_attr(tc[0], 'TableTotal', 0) or 0)
                    except Exception as tce:
                        diag['tableProbeErr'] = 'count: ' + str(tce)
                    try:
                        mx = list(q.QuerySql("SELECT FORMAT(MAX(TransactionDate), 'yyyy-MM-dd') AS LatestDate FROM [Transaction]"))
                        if mx:
                            diag['latestTxnDate'] = str(self.safe_get_attr(mx[0], 'LatestDate', '') or '')
                    except Exception as mxe:
                        if not diag.get('tableProbeErr'):
                            diag['tableProbeErr'] = 'max: ' + str(mxe)
                    # Filtered counts. Any exception here goes into
                    # diag['error'] so the UI can render it (no more
                    # silently-swallowed failures).
                    try:
                        diag_sql = """
                            SELECT
                                COUNT(*) AS TotalTxns,
                                SUM(CASE WHEN t.amt <> 0 THEN 1 ELSE 0 END) AS NonZeroAmt,
                                SUM(CASE WHEN t.[Message] LIKE 'CHK%' THEN 1 ELSE 0 END) AS ChkCount,
                                SUM(CASE WHEN t.[Message] LIKE 'CSH%' THEN 1 ELSE 0 END) AS CshCount,
                                SUM(CASE WHEN t.[Message] LIKE 'Response%' THEN 1 ELSE 0 END) AS CCCount,
                                SUM(CASE WHEN t.TransactionId LIKE 'Coupon%' THEN 1 ELSE 0 END) AS CouponCount,
                                SUM(CASE WHEN t.amt > 0
                                          AND t.[Message] NOT LIKE 'CHK%'
                                          AND t.[Message] NOT LIKE 'CSH%'
                                          AND t.[Message] NOT LIKE 'Response%'
                                          AND t.TransactionId NOT LIKE 'Coupon%'
                                          AND ISNULL(t.[Message], '') NOT LIKE 'ADJ|%'
                                          AND ISNULL(t.[Message], '') NOT LIKE 'FEE|%'
                                          AND ISNULL(t.[Message], '') NOT LIKE 'variable%'
                                         THEN 1 ELSE 0 END) AS OtherCount,
                                SUM(CASE WHEN t.[Message] LIKE '{0}%' OR t.[Message] LIKE '{1}%' THEN 1 ELSE 0 END) AS PMCount
                            FROM [Transaction] t
                            WHERE t.TransactionDate >= '{2}'
                              AND t.TransactionDate <  DATEADD(day, 1, '{3}')
                        """.format(pm_chk, pm_csh, date_from, date_to)
                        d = list(q.QuerySql(diag_sql))
                        if d:
                            diag['totalTxns']   = int(self.safe_get_attr(d[0], 'TotalTxns', 0) or 0)
                            diag['nonZeroAmt']  = int(self.safe_get_attr(d[0], 'NonZeroAmt', 0) or 0)
                            diag['chkCount']    = int(self.safe_get_attr(d[0], 'ChkCount', 0) or 0)
                            diag['cshCount']    = int(self.safe_get_attr(d[0], 'CshCount', 0) or 0)
                            diag['ccCount']     = int(self.safe_get_attr(d[0], 'CCCount', 0) or 0)
                            diag['couponCount'] = int(self.safe_get_attr(d[0], 'CouponCount', 0) or 0)
                            diag['otherCount']  = int(self.safe_get_attr(d[0], 'OtherCount', 0) or 0)
                            diag['pmCount']     = int(self.safe_get_attr(d[0], 'PMCount', 0) or 0)
                            diag['extCount']    = (diag['chkCount'] + diag['cshCount']) - diag['pmCount']
                    except Exception as diag_err:
                        diag['error'] = 'diag_sql: ' + str(diag_err)
                return self.create_json_response(True, str(len(rows)) + " receipt(s) found",
                                                {'receipts': rows, 'diag': diag, 'includeExternal': include_external})
            except Exception as e:
                return self.create_json_response(False, "Error finding receipts: " + str(e))

        def _build_receipt_html(self, row, settings):
            """Build the receipt body. Re-uses the PM_PaymentMade_Email
            template (or the bundled default) so a reprint looks identical
            to the original send. The 'previous' balance can't be
            historically reconstructed without a balance snapshot table,
            so we frame this as a 'Payment Receipt -- duplicate' and show
            current balance instead, clearly labelled."""
            try:
                tpl = model.HtmlContent(settings.get('paymentConfirmationTemplate') or PAYMENT_CONFIRMATION_TEMPLATE_NAME)
            except:
                tpl = None
            if not tpl:
                tpl = DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE
            sender = settings.get('defaultSender') or DEFAULT_EMAIL_SENDER
            charge_notes = '${:,.2f} -- {} payment on {}'.format(
                abs(float(row.get('amount') or 0)),
                row.get('paymentType', ''),
                row.get('tranDateOnly', ''))
            # Enrich charge_notes with the involvement + registrants and
            # the staff note (if any). All three appear below the amount
            # line as subtle dimmer text, so every template -- default or
            # user-custom -- benefits via the existing {chargeNotes}
            # placeholder. Templates that opt into {orgName},
            # {registrants}, or {note} directly are also handled below.
            org_name    = str(row.get('orgName', '') or '').strip()
            registrants = str(row.get('registrants', '') or '').strip()
            raw_msg     = str(row.get('message', '') or '').strip()
            # Extract the staff note from a typed Message ("CHK|note",
            # "CSH|note", "FEE|note", "ADJ|note"). Skip CC gateway noise
            # ("Response:...") and pre-piped patterns.
            staff_note = ''
            if '|' in raw_msg:
                staff_note = raw_msg.split('|', 1)[1].strip()
            elif raw_msg and not raw_msg.startswith('Response'):
                staff_note = raw_msg
            details_html = ''
            if org_name:
                details_html += ('<br><span style="color:#666;font-size:13px;">'
                                 'For: <strong>' + org_name + '</strong></span>')
            if registrants and registrants.lower() != str(row.get('person', '') or '').strip().lower():
                details_html += ('<br><span style="color:#666;font-size:13px;">'
                                 'Registered: <strong>' + registrants + '</strong></span>')
            if staff_note:
                details_html += ('<br><span style="color:#666;font-size:13px;font-style:italic;">'
                                 'Note: ' + staff_note + '</span>')
            charge_notes_full = charge_notes + details_html
            # Bulletproof balance lookup (legacy get_current_balance has
            # an except:return 0.0 that swallows IronPython attribute
            # errors -- a glitched call would print "$0 balance" on a
            # receipt that actually has money owed). Inline SUM with
            # defensive parsing instead.
            cur_balance = 0.0
            pid_v = row.get('peopleId') or 0
            oid_v = row.get('orgId') or 0
            if pid_v and oid_v:
                try:
                    bsql = ("SELECT SUM(ISNULL(IndDue, 0)) AS B "
                            "FROM dbo.TransactionSummary "
                            "WHERE PeopleId = {0} AND OrganizationId = {1}"
                            .format(int(pid_v), int(oid_v)))
                    bres = list(q.QuerySql(bsql))
                    if bres:
                        raw_b = getattr(bres[0], 'B', None)
                        if raw_b is not None:
                            cur_balance = float(raw_b)
                except:
                    cur_balance = 0.0
            body = tpl
            body = body.replace('{name}',         str(row.get('person', '')))
            body = body.replace('{chargeNotes}',  charge_notes_full)
            body = body.replace('{orgName}',      org_name)
            body = body.replace('{registrants}',  registrants)
            body = body.replace('{note}',         staff_note)
            body = body.replace('{previousDue}',  '${:,.2f}'.format(abs(float(row.get('amount') or 0)) + float(cur_balance)))
            body = body.replace('{newTotalDue}',  '${:,.2f}'.format(float(cur_balance)))
            body = body.replace('{sender_alias}', str(sender.get('sender_alias', '')))
            body = body.replace('{sender_phone}', str(sender.get('sender_phone', '')))
            body = body.replace('{sender_email}', str(sender.get('sender_email', '')))
            # Duplicate-receipt banner above the rendered template body.
            banner = ("<div style=\"background:#fff4ce;border:1px solid #f4d35e;"
                      "padding:8px 12px;border-radius:4px;font-size:12px;"
                      "color:#7a5c00;margin-bottom:12px;\">"
                      "<strong>Duplicate Receipt</strong> &mdash; original payment "
                      "recorded " + str(row.get('tranDate', '')) + ". "
                      "Balance shown reflects today's account state, not the "
                      "balance at the time of the original transaction."
                      "</div>")
            return banner + body

        def process_reprint_receipt(self):
            """Reprint mode: 'email' fires the receipt to the payer's
            email on file; 'html' returns the rendered HTML so the JS
            can pop a print window."""
            try:
                tx_id = str(getattr(model.Data, 'transaction_id', ''))
                mode = str(getattr(model.Data, 'mode', 'html')).strip().lower()
                if not tx_id:
                    return self.create_json_response(False, "Missing transaction_id")
                # Pull the single transaction back. Same column logic
                # as find_receipts (t.amt for actual amount, full
                # PaymentType CASE including CC + Coupon, registrant list).
                sql = """
                    SELECT TOP 1 t.TransactionId,
                        ISNULL(NULLIF(t.Name, ''), ISNULL(p.Name, '')) AS PersonName,
                        ISNULL(p.PeopleId, 0)     AS PeopleId,
                        ISNULL(p.EmailAddress, '') AS Email,
                        ISNULL(o.OrganizationName, '') AS OrgName,
                        ISNULL(o.OrganizationId, 0) AS OrgId,
                        t.amt                     AS Amount,
                        FORMAT(t.TransactionDate, 'yyyy-MM-dd HH:mm') AS TranDate,
                        FORMAT(t.TransactionDate, 'yyyy-MM-dd') AS TranDateOnly,
                        CASE WHEN t.[Message] LIKE 'CHK%' THEN 'Check'
                             WHEN t.[Message] LIKE 'CSH%' THEN 'Cash'
                             WHEN t.[Message] LIKE 'Response%' THEN 'Credit Card'
                             WHEN t.TransactionId LIKE 'Coupon%' THEN 'Coupon'
                             ELSE 'Other' END AS PaymentType,
                        ISNULL(t.[Message], '')   AS Message,
                        ISNULL(STUFF((
                            SELECT ', ' + p2.Name
                            FROM TransactionSummary ts2
                            JOIN People p2 ON p2.PeopleId = ts2.PeopleId
                            WHERE ts2.RegId = t.OriginalId
                              AND ts2.IsLatestTransaction = 1
                            FOR XML PATH('')
                        ), 1, 2, ''), '') AS RegistrantNames
                    FROM [Transaction] t
                    OUTER APPLY (
                        SELECT TOP 1 ts.PeopleId, ts.OrganizationId
                        FROM TransactionSummary ts
                        WHERE ts.RegId = t.OriginalId
                          AND ts.IsLatestTransaction = 1
                        ORDER BY ts.PeopleId
                    ) ts
                    LEFT JOIN People p ON p.PeopleId = ts.PeopleId
                    LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
                    WHERE t.TransactionId = '""" + tx_id.replace("'", "''") + """'
                """
                rows = list(q.QuerySql(sql))
                if not rows:
                    return self.create_json_response(False, "Transaction not found")
                r = rows[0]
                row = {
                    'transactionId': self.safe_get_attr(r, 'TransactionId', ''),
                    'person':        self.safe_get_attr(r, 'PersonName', ''),
                    'peopleId':      self.safe_get_attr(r, 'PeopleId', 0),
                    'email':         self.safe_get_attr(r, 'Email', ''),
                    'orgName':       self.safe_get_attr(r, 'OrgName', ''),
                    'orgId':         self.safe_get_attr(r, 'OrgId', 0),
                    'amount':        float(self.safe_get_attr(r, 'Amount', 0) or 0),
                    'tranDate':      self.safe_get_attr(r, 'TranDate', ''),
                    'tranDateOnly':  self.safe_get_attr(r, 'TranDateOnly', ''),
                    'paymentType':   self.safe_get_attr(r, 'PaymentType', ''),
                    'registrants':   self.safe_get_attr(r, 'RegistrantNames', ''),
                    'message':       self.safe_get_attr(r, 'Message', ''),
                }
                settings = pm_load_settings()
                body = self._build_receipt_html(row, settings)
                if mode == 'email':
                    # Defense-in-depth: refuse to email a "duplicate
                    # receipt" for a payment that didn't go through PM
                    # to begin with (no original receipt was ever sent).
                    settings = pm_load_settings()
                    pm_chk = settings.get('checkPaymentType') or CHECK_PAYMENT_TYPE
                    pm_csh = settings.get('cashPaymentType')  or CASH_PAYMENT_TYPE
                    msg = str(self.safe_get_attr(r, 'Message', '') or '')
                    if not (msg.startswith(pm_chk) or msg.startswith(pm_csh)):
                        return self.create_json_response(False,
                            "This payment was not recorded through Payment Manager. "
                            "No original receipt was sent, so a duplicate cannot be emailed. "
                            "You can still Print a receipt for the church files.")
                    if not row.get('email'):
                        return self.create_json_response(False, "No email on file for " + str(row.get('person', '')))
                    sender = settings.get('defaultSender') or DEFAULT_EMAIL_SENDER
                    subject = "Receipt (duplicate) -- " + str(row.get('paymentType', '')) + " payment on " + str(row.get('tranDateOnly', ''))
                    try:
                        model.Email(
                            int(row.get('peopleId') or 0),
                            int(sender.get('sender_id') or 0),
                            str(sender.get('sender_email', '')),
                            str(sender.get('sender_alias', '')),
                            subject,
                            body
                        )
                        return self.create_json_response(True, "Receipt emailed to " + str(row.get('email')))
                    except Exception as e:
                        return self.create_json_response(False, "Email failed: " + str(e))
                # mode == 'html' -- return for popup print.
                # mode == 'customize' -- same body but also returns the
                # subject + recipient email so the JS can pre-fill the
                # Customize modal.
                subject = 'Receipt (duplicate) -- ' + str(row.get('paymentType', '')) + ' payment on ' + str(row.get('tranDateOnly', ''))
                return self.create_json_response(True, "ok", {
                    'html': body,
                    'subject': subject,
                    'person': row.get('person', ''),
                    'email': row.get('email', ''),
                    'peopleId': row.get('peopleId', 0),
                    'paymentType': row.get('paymentType', ''),
                    'tranDateOnly': row.get('tranDateOnly', ''),
                })
            except Exception as e:
                return self.create_json_response(False, "Reprint error: " + str(e))

        def process_send_custom_receipt(self):
            """Send a hand-edited receipt body. Lets a staffer add custom
            language ("thanks for your generous gift", "see you Sunday",
            etc.) on top of the rendered receipt before firing it. The
            same external-payment guard as reprint applies: we will not
            send for payments that didn't go through PM.

            Inputs (URL-encoded body to dodge ASP.NET validation):
              transaction_id, subject, body, encoded='urlc'
            """
            try:
                tx_id    = str(getattr(model.Data, 'transaction_id', '') or '').strip()
                subject  = str(getattr(model.Data, 'subject', '') or '')
                body     = str(getattr(model.Data, 'body', '') or '')
                to_email = str(getattr(model.Data, 'to_email', '') or '').strip()
                if str(getattr(model.Data, 'encoded', '') or '') == 'urlc':
                    try:
                        import urllib
                        subject = urllib.unquote(subject)
                        body    = urllib.unquote(body)
                        try: subject = subject.decode('utf-8')
                        except: pass
                        try: body = body.decode('utf-8')
                        except: pass
                    except:
                        pass
                if not tx_id:    return self.create_json_response(False, "Missing transaction_id")
                if not subject:  return self.create_json_response(False, "Subject is required")
                if not body:     return self.create_json_response(False, "Receipt body is empty")
                if not to_email: return self.create_json_response(False, "Recipient email is required")
                # Server-side email sanity check (JS validates too).
                import re as _re
                if not _re.match(r'^[^\s@]+@[^\s@]+\.[^\s@]+$', to_email):
                    return self.create_json_response(False, "Invalid recipient email address")
                # Pull the txn back so we know who to send to + can
                # enforce the PM-only guard.
                sql = """
                    SELECT TOP 1 t.TransactionId,
                        ISNULL(NULLIF(t.Name, ''), ISNULL(p.Name, '')) AS PersonName,
                        ISNULL(p.PeopleId, 0)     AS PeopleId,
                        ISNULL(p.EmailAddress, '') AS Email,
                        ISNULL(t.[Message], '')   AS Message
                    FROM [Transaction] t
                    LEFT JOIN TransactionSummary ts ON ts.RegId = t.OriginalId AND ts.IsLatestTransaction = 1
                    LEFT JOIN People p ON p.PeopleId = ts.PeopleId
                    WHERE t.TransactionId = '""" + tx_id.replace("'", "''") + """'
                """
                rows = list(q.QuerySql(sql))
                if not rows:
                    return self.create_json_response(False, "Transaction not found")
                r = rows[0]
                people_id = int(self.safe_get_attr(r, 'PeopleId', 0) or 0)
                email     = str(self.safe_get_attr(r, 'Email', '') or '')
                msg       = str(self.safe_get_attr(r, 'Message', '') or '')
                person    = str(self.safe_get_attr(r, 'PersonName', '') or '')
                settings = pm_load_settings()
                # No PM-origin guard here -- Customize generates a fresh
                # receipt the staffer authored, not a duplicate of a
                # prior PM send. It's appropriate for any payment type
                # (CC, Coupon, external CHK/CSH, etc.).
                sender = settings.get('defaultSender') or DEFAULT_EMAIL_SENDER
                is_override = (to_email.lower() != (email or '').lower())
                # The model.Email path needs a real people_id. When we
                # don't have one (external rows where ts join returned
                # NULL, even if the staffer kept the on-file email),
                # fall through to SendEmail too.
                use_send_email = is_override or not people_id
                try:
                    if use_send_email:
                        # SendEmail with an explicit recipient. We try to
                        # find an existing person by email first to avoid
                        # accidentally creating a stub record; if there's
                        # no match, fall back to passing the email string
                        # directly (TouchPoint will route it).
                        alt_person = None
                        try:
                            existing_pid = model.FindPersonId(to_email)
                            if existing_pid:
                                alt_person = model.GetPerson(int(existing_pid))
                        except:
                            alt_person = None
                        try:
                            if alt_person:
                                model.SendEmail(
                                    alt_person,
                                    str(sender.get('sender_email', '')),
                                    str(sender.get('sender_alias', '')),
                                    subject,
                                    body
                                )
                            else:
                                model.SendEmail(
                                    to_email,
                                    str(sender.get('sender_email', '')),
                                    str(sender.get('sender_alias', '')),
                                    subject,
                                    body
                                )
                        except Exception as se:
                            return self.create_json_response(False, "SendEmail failed: " + str(se))
                        if is_override and email:
                            return self.create_json_response(True, "Custom receipt emailed to " + to_email + " (override -- email on file was " + email + ")")
                        return self.create_json_response(True, "Custom receipt emailed to " + to_email)
                    else:
                        # Standard path: PM payment with a known person.
                        model.Email(
                            people_id,
                            int(sender.get('sender_id') or 0),
                            str(sender.get('sender_email', '')),
                            str(sender.get('sender_alias', '')),
                            subject,
                            body
                        )
                        return self.create_json_response(True, "Custom receipt emailed to " + email)
                except Exception as e:
                    return self.create_json_response(False, "Email failed: " + str(e))
            except Exception as e:
                return self.create_json_response(False, "Send custom error: " + str(e))

        # ==============================================================
        # SETTINGS -- save/load via UI
        # ==============================================================
        def process_save_settings(self):
            try:
                raw = str(getattr(model.Data, 'settings_json', '') or '')
                if not raw:
                    return self.create_json_response(False, "Missing settings payload")
                # ASP.NET request validation may have rejected HTML in
                # the senders dict; the client URL-encodes when there's
                # any HTML risk.
                if str(getattr(model.Data, 'encoded', '') or '') == 'urlc':
                    try:
                        import urllib
                        raw = urllib.unquote(raw)
                        try:
                            raw = raw.decode('utf-8')
                        except:
                            pass
                    except:
                        pass
                data = json.loads(raw)
                if not isinstance(data, dict):
                    return self.create_json_response(False, "Settings payload not a JSON object")
                data = _pm_normalize_settings(data)
                if not pm_save_settings(data):
                    return self.create_json_response(False, "Could not write settings to TextContent")
                return self.create_json_response(True, "Settings saved", {'settings': data})
            except Exception as e:
                return self.create_json_response(False, "Save settings failed: " + str(e))

        def process_load_settings(self):
            try:
                s = pm_load_settings()
                # Also pull the two HTML templates so the editor in the
                # Settings modal can show what's actually rendering. Falls
                # back to the built-in defaults so the editor never opens
                # empty just because no custom HtmlContent exists yet.
                payload = {'settings': s, 'templates': {}}
                for key, name, fallback in [
                    ('confirmation', s.get('paymentConfirmationTemplate') or PAYMENT_CONFIRMATION_TEMPLATE_NAME, DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE),
                    ('notification', s.get('paymentNotificationTemplate') or PAYMENT_NOTIFICATION_TEMPLATE_NAME, DEFAULT_PAYMENT_EMAIL_TEMPLATE),
                ]:
                    body = None
                    is_default = True
                    try:
                        body = model.HtmlContent(name)
                    except:
                        body = None
                    # Treat None / empty / whitespace-only as "not there".
                    # An admin who created the HtmlContent record but never
                    # filled it in shouldn't get a blank editor.
                    if not body or not str(body).strip():
                        body = fallback
                    else:
                        is_default = False
                    payload['templates'][key] = {
                        'name': name,
                        'body': body,
                        'isDefault': is_default,
                    }
                return self.create_json_response(True, "ok", payload)
            except Exception as e:
                return self.create_json_response(False, "Load settings failed: " + str(e))

        def process_save_template(self):
            """Save one of the two named templates to HtmlContent."""
            try:
                which = str(getattr(model.Data, 'which', '')).strip().lower()
                body  = str(getattr(model.Data, 'body', '') or '')
                if str(getattr(model.Data, 'encoded', '') or '') == 'urlc':
                    try:
                        import urllib
                        body = urllib.unquote(body)
                        try:
                            body = body.decode('utf-8')
                        except:
                            pass
                    except:
                        pass
                if which not in ('confirmation', 'notification'):
                    return self.create_json_response(False, "Unknown template key: " + which)
                s = pm_load_settings()
                if which == 'confirmation':
                    name = s.get('paymentConfirmationTemplate') or PAYMENT_CONFIRMATION_TEMPLATE_NAME
                else:
                    name = s.get('paymentNotificationTemplate') or PAYMENT_NOTIFICATION_TEMPLATE_NAME
                try:
                    model.WriteContentHtml(name, body, '')
                    return self.create_json_response(True, 'Saved template "' + name + '"', {'name': name})
                except Exception as we:
                    return self.create_json_response(False, "Write failed: " + str(we))
            except Exception as e:
                return self.create_json_response(False, "Save template failed: " + str(e))

        def process_reset_template(self):
            """Reset one template to its built-in default body."""
            try:
                which = str(getattr(model.Data, 'which', '')).strip().lower()
                if which not in ('confirmation', 'notification'):
                    return self.create_json_response(False, "Unknown template key: " + which)
                default_body = (DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE
                                if which == 'confirmation'
                                else DEFAULT_PAYMENT_EMAIL_TEMPLATE)
                return self.create_json_response(True, 'ok', {'body': default_body})
            except Exception as e:
                return self.create_json_response(False, "Reset failed: " + str(e))

        # ==============================================================
        # AUTO-UPDATE
        # ==============================================================
        def process_apply_update(self):
            try:
                target_name = pm_get_script_name() or DC_SCRIPT_ID
                fetch_url = DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID
                # 1. Pull the new code. RestGet can return None, an
                #    HTML error page, or a JSON error envelope when the
                #    worker fails. Sanity-check before touching anything.
                try:
                    new_code = model.RestGet(fetch_url, {})
                except Exception as fe:
                    return self.create_json_response(False,
                        "Couldn't reach the update server. " + str(fe) +
                        " | Manual install: " + fetch_url)
                if new_code is None:
                    return self.create_json_response(False,
                        "Update server returned nothing. Try again in a minute or use the manual link: " + fetch_url)
                new_code = str(new_code)
                if len(new_code) < 500:
                    return self.create_json_response(False,
                        "Update payload was too small to be a real script (" + str(len(new_code)) +
                        " bytes). Try again or use the manual link: " + fetch_url)
                # Look for the unmistakable header so an HTML error page
                # or a JSON envelope can't slip through.
                if 'TPxi_PaymentManager' not in new_code or 'APP_VERSION' not in new_code:
                    return self.create_json_response(False,
                        "Update payload didn't look like Payment Manager code. The update server may be returning an error page. "
                        "Manual install: " + fetch_url)
                # 2. Write to Special Content. Verified against the
                #    bvcms source: WriteContentPython has no role check,
                #    so a write failure here is either a TouchPoint-side
                #    SQL / encoding error or the script name we resolved
                #    doesn't match the actual installed script.
                try:
                    model.WriteContentPython(target_name, new_code)
                except Exception as we:
                    return self.create_json_response(False,
                        "Write failed: " + str(we) +
                        " | If this persists, paste the latest code from " + fetch_url +
                        " into Admin > Advanced > Special Content > Python > " + target_name + ".")
                # 3. Verify the write actually landed at the right name
                #    by reading it back and checking for the new version
                #    marker. If the script is installed under a name we
                #    couldn't auto-detect, the write went to the wrong
                #    place and the running script is unchanged.
                try:
                    written = model.PythonContent(target_name) or ''
                    if APP_VERSION not in str(written):
                        return self.create_json_response(False,
                            "Update wrote to '" + target_name + "' but the running version still shows " +
                            APP_VERSION + ". This usually means the script is installed under a different "
                            "name. Find your installed script under Admin > Advanced > Special Content > "
                            "Python and paste the latest code from " + fetch_url + " into it directly.")
                except:
                    # Verification is best-effort -- if PythonContent fails
                    # we trust the write succeeded.
                    pass
                return self.create_json_response(True,
                    "Updated " + target_name + " to v" + APP_VERSION +
                    ". Reload the page (hard refresh with Ctrl+Shift+R or Cmd+Shift+R if the version doesn't change).")
            except Exception as e:
                return self.create_json_response(False, "Update failed: " + str(e))

        def get_current_balance(self, payer_id, org_id):
            """Get current balance for payer/org"""
            try:
                sql = "SELECT Sum(IndDue) as IndDue FROM dbo.TransactionSummary WHERE PeopleId = {0} AND OrganizationId = {1}".format(payer_id, org_id)
                result = q.QuerySql(sql)
                if result and len(result) > 0:
                    balance = self.safe_get_attr(result[0], 'IndDue', 0)
                    return float(balance) if balance is not None else 0.0
                return 0.0
            except:
                return 0.0

        def build_payment_email(self, name, charge_amount, previous_due, total_due, paylink, email_details):
            """Build payment notification email"""
            try:
                message = model.HtmlContent(PAYMENT_NOTIFICATION_TEMPLATE_NAME)
            except:
                message = DEFAULT_PAYMENT_EMAIL_TEMPLATE
            
            charge_notes = '${:,.2f}........New Charge'.format(float(charge_amount))
            previous_due_text = '${:,.2f}........Previous Balance'.format(float(previous_due))
            total_due_text = '${:,.2f}'.format(float(total_due))
            
            return message.format(
                name=name,
                chargeNotes=charge_notes,
                previousDue=previous_due_text,
                totalDue=total_due_text,
                paylink=paylink,
                sender_alias=email_details.get('sender_alias', 'Payment System'),
                sender_phone=email_details.get('sender_phone', ''),
                sender_email=email_details.get('sender_email', '')
            ) + '{track}{tracklinks}<br />'

        def build_payment_confirmation_email(self, name, payment_amount, previous_due, new_total, payment_type, payment_desc, email_details):
            """Build payment confirmation email"""
            try:
                # Try to load the template from TouchPoint's special content
                message = model.HtmlContent(PAYMENT_CONFIRMATION_TEMPLATE_NAME)
            except Exception as e:
                # If template not found, log error and use default template
                print("<!-- Warning: Email template not found. Using default. Error: {} -->".format(str(e)))
                message = DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE
            
            # Format payment details
            payment_amount_abs = abs(float(payment_amount))
            charge_notes = '{0} ${1:.2f}: {2}'.format(
                payment_type.replace('|', ''), 
                payment_amount_abs, 
                payment_desc
            )
            
            # Format monetary values
            previous_due_text = '${:,.2f}'.format(float(previous_due))
            new_total_text = '${:,.2f}'.format(float(new_total))
            
            # Apply template replacements
            try:
                return message.format(
                    name=name,
                    chargeNotes=charge_notes,
                    previousDue=previous_due_text,
                    newTotalDue=new_total_text,
                    sender_alias=email_details.get('sender_alias', 'Payment System'),
                    sender_phone=email_details.get('sender_phone', ''),
                    sender_email=email_details.get('sender_email', '')
                )
            except Exception as e:
                # If template format fails, return simplified message
                print("<!-- Warning: Error formatting email template: {} -->".format(str(e)))
                return """
                <p>Dear {0},</p>
                <p>We have received your payment of {1}. Thank you!</p>
                <p>New Balance: {2}</p>
                <p>Thank you,<br>{3}</p>
                """.format(name, charge_notes, new_total_text, email_details.get('sender_alias', 'Payment System'))
            
            return message.format(
                name=name,
                chargeNotes=charge_notes,
                previousDue=previous_due_text,
                newTotalDue=new_total_text,
                sender_alias=email_details.get('sender_alias', 'Payment System'),
                sender_phone=email_details.get('sender_phone', ''),
                sender_email=email_details.get('sender_email', '')
            )

        def create_json_response(self, success, message, data=None):
            """Create JSON response for AJAX calls.

            The 'data' dict (if provided) is *merged into the root* of
            the response so that JS handlers can read fields directly
            (e.g. d.settings, d.receipts, d.diag). Earlier versions
            nested data under response['data'], which silently broke
            every JS callsite that expected a flat object -- diagnostic
            zeros were just JS undefined-fallbacks, not real DB counts.
            """
            response = {
                'success': success,
                'message': message
            }
            if data:
                for k, v in data.items():
                    # Don't let payload keys clobber 'success' / 'message'.
                    if k not in response:
                        response[k] = v
            return json.dumps(response)

        # View rendering methods
        def render_programs_view(self):
            """Render the main programs overview"""
            programs = self.get_programs_with_dues()
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-credit-card"></i> Payment Manager &mdash; Outstanding Balances</h3>
                    <div class="pm-actions">
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>

                <div class="pm-content">
                    <table class="pm-table">
                        <thead>
                            <tr>
                                <th>Program Name</th>
                                <th>Outstanding Amount</th>
                                <th>Payers Count</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
            """
            
            total_outstanding = 0
            total_payers = 0
            
            try:
                for program in programs:
                    outstanding = float(self.safe_get_attr(program, 'Outstanding', 0) or 0)
                    payer_count = int(self.safe_get_attr(program, 'PayerCount', 0) or 0)
                    total_outstanding += outstanding
                    total_payers += payer_count
                    
                    program_name = self.safe_get_attr(program, 'ProgramName', 'Unknown')
                    program_id = self.safe_get_attr(program, 'ProgramId', 0)
                    
                    html += """
                                <tr>
                                    <td><strong>{0}</strong></td>
                                    <td class="pm-currency">{1}</td>
                                    <td class="pm-center">{2}</td>
                                    <td class="pm-center">
                                        <button class="btn btn-sm btn-outline-primary" 
                                                onclick="viewDivisions({3})">
                                            <i class="fa fa-list"></i> View Divisions
                                        </button>
                                        <button class="btn btn-sm btn-outline-success" 
                                                onclick="viewPayers(null, {3})">
                                            <i class="fa fa-users"></i> View All Payers
                                        </button>
                                    </td>
                                </tr>
                    """.format(
                        program_name,
                        self.format_currency(outstanding),
                        payer_count,
                        program_id
                    )
            except Exception as e:
                html += "<tr><td colspan='4'>Error loading programs: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                        <tfoot>
                            <tr class="pm-total">
                                <th style="text-align: left;">Total</th>
                                <th class="pm-currency" style="text-align: right;">{0}</th>
                                <th class="pm-center" style="text-align: center;">{1}</th>
                                <th style="text-align: center;"></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            """.format(
                self.format_currency(total_outstanding),
                total_payers
            )
            
            return html
            
        def render_divisions_view(self, program_id):
            """Render divisions/involvements within a program"""
            divisions = self.get_divisions_with_dues(program_id)
            program_name = "Unknown Program"
            
            # Get program name
            try:
                program_data = q.QuerySql("SELECT Name FROM Program WHERE Id = {0}".format(program_id))
                if program_data and len(program_data) > 0:
                    program_name = self.safe_get_attr(program_data[0], 'Name', 'Unknown Program')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-sitemap"></i> {0} - Divisions & Involvements</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="viewPrograms()">
                            <i class="fa fa-arrow-left"></i> Back to Outstanding Balances
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-content">
                    <table class="pm-table">
                        <thead>
                            <tr>
                                <th>Division</th>
                                <th>Involvement</th>
                                <th>Outstanding</th>
                                <th>Payers</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(program_name)
            
            current_division = ""
            total_outstanding = 0
            total_payers = 0
            
            try:
                for div in divisions:
                    outstanding = float(self.safe_get_attr(div, 'Outstanding', 0) or 0)
                    payer_count = int(self.safe_get_attr(div, 'PayerCount', 0) or 0)
                    total_outstanding += outstanding
                    total_payers += payer_count
                    
                    division_name = self.safe_get_attr(div, 'DivisionName', '')
                    org_name = self.safe_get_attr(div, 'OrganizationName', 'Unknown')
                    org_id = self.safe_get_attr(div, 'OrganizationId', 0)
                    
                    division_cell = ""
                    if division_name != current_division:
                        division_cell = division_name
                        current_division = division_name
                    
                    html += """
                                <tr>
                                    <td>{0}</td>
                                    <td>{1}</td>
                                    <td class="pm-currency">{2}</td>
                                    <td class="pm-center">{3}</td>
                                    <td class="pm-center">
                                        <button class="btn btn-sm btn-outline-success" 
                                                onclick="viewPayers({4}, {5})">
                                            <i class="fa fa-users"></i> View Payers
                                        </button>
                                    </td>
                                </tr>
                    """.format(
                        division_cell,
                        org_name,
                        self.format_currency(outstanding),
                        payer_count,
                        org_id,
                        program_id
                    )
            except Exception as e:
                html += "<tr><td colspan='5'>Error loading divisions: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                        <tfoot>
                            <tr class="pm-total">
                                <th colspan="2">Total</th>
                                <th class="pm-currency">{0}</th>
                                <th class="pm-center">{1}</th>
                                <th></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            """.format(
                self.format_currency(total_outstanding),
                total_payers
            )
            
            return html
            
        def render_payers_view(self, org_id=None, program_id=None):
            """Render individual payers with payment options"""
            payers = self.get_payers_with_dues(org_id, program_id)
            
            # Get context title
            context_title = "All Programs"
            try:
                if program_id:
                    program_data = q.QuerySql("SELECT Name FROM Program WHERE Id = {0}".format(program_id))
                    if program_data and len(program_data) > 0:
                        context_title = self.safe_get_attr(program_data[0], 'Name', 'Unknown Program')
                if org_id:
                    org_data = q.QuerySql("SELECT OrganizationName FROM Organizations WHERE OrganizationId = {0}".format(org_id))
                    if org_data and len(org_data) > 0:
                        context_title = self.safe_get_attr(org_data[0], 'OrganizationName', 'Unknown Organization')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-users"></i> {0} - Individual Payers</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="history.go(-1)">
                            <i class="fa fa-arrow-left"></i> Go Back
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-search">
                    <input type="text" id="searchInput" placeholder="Search by name, email, or phone..." 
                           class="form-control">
                </div>
                
                <div class="pm-content">
                    <table class="pm-table" id="payersTable">
                        <thead>
                            <tr>
                                <th>Name & Contact</th>
                                <th>Involvement</th>
                                <th>Outstanding</th>
                                <th>Payment Options</th>
                                <th>History</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(context_title)
            
            total_outstanding = 0
            
            try:
                for payer in payers:
                    outstanding = float(self.safe_get_attr(payer, 'Outstanding', 0) or 0)
                    total_outstanding += outstanding
                    is_refund_driven = int(self.safe_get_attr(payer, 'IsRefundDriven', 0) or 0) == 1
                    coregistrant_names = str(self.safe_get_attr(payer, 'CoRegistrantNames', '') or '').strip()

                    # Get payer details safely
                    payer_name = self.safe_get_attr(payer, 'Name2', 'Unknown')
                    payer_id = self.safe_get_attr(payer, 'PeopleId', 0)
                    payer_email = self.safe_get_attr(payer, 'EmailAddress', '')
                    org_name = self.safe_get_attr(payer, 'OrganizationName', 'Unknown')
                    org_id = self.safe_get_attr(payer, 'OrganizationId', 0)
                    division_name = self.safe_get_attr(payer, 'Division', '')
                    family_id = self.safe_get_attr(payer, 'FamilyId', 0)
                    
                    # Get parent contact info
                    cc_emails, parent_info = self.get_parent_emails(family_id)
                    
                    # Format contact information
                    phone_info = ""
                    cell_phone = self.safe_get_attr(payer, 'CellPhone', '')
                    home_phone = self.safe_get_attr(payer, 'HomePhone', '')
                    if cell_phone:
                        phone_info += " <small>C: {0}</small>".format(self.format_phone(cell_phone))
                    if home_phone:
                        phone_info += " <small>H: {0}</small>".format(self.format_phone(home_phone))
                    
                    email_display = payer_email if payer_email else "<em>No email</em>"
                    if parent_info and parent_info.get('email'):
                        email_display += "<br><small>Parent: {0}</small>".format(parent_info['email'])
                    
                    # Payment buttons -- separate Email vs Text + Record
                    # + Adjust. Adjust is offered on every row (even when
                    # outstanding == 0) so credits can be applied
                    # proactively without a balance trigger.
                    payment_buttons = ""
                    if outstanding > 0:
                        payment_buttons = """
                            <div class="pm-btn-group" style="position:relative;">
                                <button class="btn btn-sm btn-outline-primary"
                                        onclick="sendPaymentLink({0}, {1}, '{2}', {3}, '{4}', 'email')"
                                        title="Email the payment link to the person">
                                    <i class="fa fa-envelope"></i> Email Link
                                </button>
                                <button class="btn btn-sm btn-outline-primary"
                                        onclick="sendPaymentLink({0}, {1}, '{2}', {3}, '{4}', 'text')"
                                        title="Text the payment link to the person's cell">
                                    <i class="fa fa-mobile"></i> Text Link
                                </button>
                                <button class="btn btn-sm btn-outline-success"
                                        onclick="recordPayment({0}, {1}, '{2}', {3})">
                                    <i class="fa fa-money"></i> Record Payment
                                </button>
                                <button class="btn btn-sm btn-outline-secondary"
                                        onclick="openAdjustBalance({0}, {1}, '{2}', {3})"
                                        title="Add a charge or apply a credit to this balance">
                                    <i class="fa fa-pencil"></i> Adjust
                                </button>
                            </div>
                        """.format(
                            payer_id,
                            org_id,
                            payer_name.replace("'", "\\'"),
                            outstanding,
                            cc_emails
                        )

                    # Refund-driven row: badge + one-click Zero Out.
                    # When co-registrants share the same RegId (e.g. a
                    # family signed up together and got refunded), show
                    # their names so staff can SEE the connection, and
                    # update the Zero Out button to indicate "N people"
                    # so they know it will clear everyone in one shot.
                    refund_badge_html = ''
                    if is_refund_driven and outstanding > 0:
                        esc_name = payer_name.replace("'", "\\'")
                        coreg_list = [n.strip() for n in coregistrant_names.split(',') if n.strip()]
                        coreg_count = len(coreg_list)
                        # Visual link to co-registrants -- amber chip
                        # under the refund pill listing the names.
                        coreg_chip = ''
                        if coreg_count > 0:
                            others_str = ', '.join(coreg_list)
                            coreg_chip = (
                                '<div style="margin-top:3px;display:inline-flex;align-items:center;gap:4px;'
                                'background:#fff4ce;color:#7a5c00;padding:2px 8px;border-radius:10px;'
                                'font-size:10px;font-weight:600;cursor:help;border:1px solid #f4d35e;" '
                                'title="This refund was for a multi-person registration. Zero Out will clear '
                                'all of them together so no one is left dangling.">'
                                '<i class="fa fa-link"></i> linked: ' + others_str +
                                '</div>'
                            )
                        # Zero Out button label reflects multi-person
                        # so staff aren't surprised when their click
                        # clears co-registrants too.
                        if coreg_count > 0:
                            btn_label = 'Zero Out (' + str(coreg_count + 1) + ' people)'
                            btn_title = ('Apply a credit to ALL ' + str(coreg_count + 1) +
                                         ' people on this registration (' + payer_name + ' + ' +
                                         ', '.join(coreg_list) + ') so the whole group disappears from '
                                         'Outstanding Balances together. Audit trail preserved per person.')
                        else:
                            btn_label = 'Zero Out'
                            btn_title = ('Apply a credit equal to the outstanding amount so this person '
                                         'no longer appears in Outstanding Balances. Posts an '
                                         'ADJ|Refund-cleared adjustment for audit.')
                        # Build the onclick string cleanly. JS signature:
                        # pmZeroOutRefund(payerId, orgId, payerName,
                        # outstanding, coregCount, coregNamesStr).
                        if coreg_count > 0:
                            safe_others = others_str.replace("\\", "\\\\").replace("'", "\\'")
                        else:
                            safe_others = ''
                        onclick_call = (
                            "pmZeroOutRefund(" +
                            str(payer_id) + ", " +
                            str(org_id or 0) + ", " +
                            "'" + esc_name + "', " +
                            str(outstanding) + ", " +
                            str(coreg_count) + ", " +
                            "'" + safe_others + "'" +
                            ")"
                        )
                        refund_badge_html = (
                            '<div style="margin-top:3px;display:inline-flex;align-items:center;gap:4px;'
                            'background:#f8d7da;color:#a8071a;padding:2px 8px;border-radius:10px;'
                            'font-size:10px;font-weight:700;cursor:help;" '
                            'title="The most recent activity on this registration was a refund matching the registration total. '
                            'The person(s) were refunded -- they do not actually owe. Use Zero Out to clear the balance.">'
                            '<i class="fa fa-undo"></i> refund issued'
                            '</div>'
                            + coreg_chip +
                            '<div style="margin-top:4px;">'
                            '<button class="btn btn-sm" '
                            'onclick="' + onclick_call + '" '
                            'style="padding:2px 8px;font-size:11px;background:#a8071a;color:#fff;'
                            'border:none;border-radius:4px;cursor:pointer;font-weight:600;" '
                            'title="' + btn_title + '">'
                            '<i class="fa fa-eraser"></i> ' + btn_label +
                            '</button></div>'
                        )

                    html += """
                                <tr>
                                    <td>
                                        <div>
                                            <strong>{0}</strong>
                                            <a href="{1}/Person2/{2}" target="_blank">
                                                <i class="fa fa-external-link"></i>
                                            </a>
                                        </div>
                                        <div class="pm-contact">
                                            {3}<br>
                                            {4}
                                        </div>
                                    </td>
                                    <td>
                                        <div>{5}</div>
                                        <small class="pm-muted">{6}</small>
                                    </td>
                                    <td class="pm-currency">
                                        <a href="javascript:void(0)"
                                           onclick="pmToggleChargeDetails(this, {2}, {9})"
                                           style="color:inherit;text-decoration:none;display:inline-flex;align-items:center;gap:6px;cursor:pointer;"
                                           title="Toggle payment history for this involvement -- charges, payments, adjustments, and running balance">
                                            <i class="fa fa-caret-right pm-charge-toggle" style="color:#666;transition:transform 0.15s;"></i>
                                            <span>{7}</span>
                                        </a>
                                        {10}
                                        <div style="font-size:10px;color:#999;margin-top:2px;line-height:1;">payment history &raquo;</div>
                                    </td>
                                    <td class="pm-center">
                                        {8}
                                    </td>
                                    <td class="pm-center">
                                        <div class="pm-btn-group">
                                            <button class="btn btn-sm btn-outline-info"
                                                    onclick="viewTransactions({2})"
                                                    title="Full transaction history across ALL involvements (separate page). For just this involvement, click the caret next to Outstanding for the inline payment history.">
                                                <i class="fa fa-history"></i> All-Org History
                                            </button>
                                            <button class="btn btn-sm btn-outline-secondary"
                                                    onclick="viewEmails({2})">
                                                <i class="fa fa-envelope"></i> Email History
                                            </button>
                                        </div>
                                    </td>
                                </tr>
                                <tr class="pm-charge-row" data-payer="{2}" data-org="{9}" style="display:none;background:#f8f9fa;">
                                    <td colspan="5" style="padding:10px 18px;">
                                        <div class="pm-charge-body" style="font-size:12px;color:#555;">Loading charge history...</div>
                                    </td>
                                </tr>
                    """.format(
                        payer_name,
                        model.CmsHost,
                        payer_id,
                        email_display,
                        phone_info,
                        org_name,
                        division_name,
                        self.format_currency(outstanding),
                        payment_buttons,
                        org_id or 0,
                        refund_badge_html
                    )
            except Exception as e:
                html += "<tr><td colspan='5'>Error loading payers: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                        <tfoot>
                            <tr class="pm-total">
                                <th colspan="2">Total Outstanding</th>
                                <th class="pm-currency">{0}</th>
                                <th colspan="2"></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            """.format(self.format_currency(total_outstanding))
            
            return html

        def render_email_history(self, people_id):
            """Render email history for a person"""
            emails = self.get_email_history(people_id)
            
            # Get person name
            person_name = "Unknown Person"
            try:
                person_data = q.QuerySql("SELECT Name FROM People WHERE PeopleId = {0}".format(people_id))
                if person_data and len(person_data) > 0:
                    person_name = self.safe_get_attr(person_data[0], 'Name', 'Unknown Person')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-envelope"></i> Email History - {0}</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="history.go(-1)">
                            <i class="fa fa-arrow-left"></i> Go Back
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-content">
                    <table class="pm-table">
                        <thead>
                            <tr>
                                <th>Sent Date</th>
                                <th>From</th>
                                <th>Subject</th>
                                <th>Opened</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(person_name)
            
            try:
                if not emails:
                    html += "<tr><td colspan='5' class='pm-center'>No email history found</td></tr>"
                else:
                    for email in emails:
                        sent_date = self.safe_get_attr(email, 'Sent', '')
                        from_name = self.safe_get_attr(email, 'FromName', 'Unknown')
                        subject = self.safe_get_attr(email, 'Subject', 'No Subject')
                        opened = self.safe_get_attr(email, 'Opened', 0)
                        message_id = self.safe_get_attr(email, 'messageId', '')
                        
                        html += """
                                    <tr>
                                        <td>{0}</td>
                                        <td>{1}</td>
                                        <td>{2}</td>
                                        <td class="pm-center">{3}</td>
                                        <td class="pm-center">
                                            <button class="btn btn-sm btn-outline-primary" 
                                                    onclick="previewEmail({4}, {5})">
                                                <i class="fa fa-eye"></i> Preview
                                            </button>
                                            <button class="btn btn-sm btn-outline-success" 
                                                    onclick="resendEmail({4}, {5})">
                                                <i class="fa fa-paper-plane"></i> Resend
                                            </button>
                                        </td>
                                    </tr>
                        """.format(
                            sent_date,
                            from_name,
                            subject,
                            opened,
                            message_id,
                            people_id
                        )
            except Exception as e:
                html += "<tr><td colspan='5'>Error loading email history: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                    </table>
                </div>
            </div>
            """
            
            return html

        def render_email_preview(self, message_id, people_id):
            """Render email preview"""
            try:
                sql = '''
                Select eq.Subject
                    ,eq.Body
                    ,eq.Sent
                    ,eq.FromName
                    ,eq.Id AS [messageId]
                    ,eqt.PeopleId
                    ,eq.QueuedBy
                    ,eq.FromAddr
                from EmailQueue eq
                left join dbo.EmailQueueTo eqt on (eqt.id = eq.id and eqt.PeopleId = {1})
                where eq.Id = {0}
                '''.format(message_id, people_id)
                
                # Use the same pattern as your original PM_EmailPreview
                emailTitle = ""
                emailBody = ""
                
                email_data = q.QuerySql(sql)
                if not email_data:
                    return "<div class='alert alert-warning'>Email not found</div>"
                
                # Use for loop like your original code
                for a in email_data:
                    return """
                    <div class="pm-container">
                        <div class="pm-header">
                            <h3><i class="fa fa-envelope-open"></i> Email Preview</h3>
                            <div class="pm-actions">
                                <button class="btn btn-secondary" onclick="history.go(-1)">
                                    <i class="fa fa-arrow-left"></i> Go Back
                                </button>
                                <button class="btn btn-success" onclick="resendEmail({0}, {1})">
                                    <i class="fa fa-paper-plane"></i> Send Copy
                                </button>
                                <button class="btn btn-primary" onclick="window.print()">
                                    <i class="fa fa-print"></i> Print
                                </button>
                            </div>
                        </div>
                        
                        <div class="pm-content">
                            <div class="email-meta">
                                <h4>Email Details</h4>
                                <table class="pm-table">
                                    <tr><th width="120">Subject:</th><td>{2}</td></tr>
                                    <tr><th>From:</th><td>{3}</td></tr>
                                    <tr><th>Sent:</th><td>{4}</td></tr>
                                </table>
                            </div>
                            
                            <div class="email-body">
                                <h4>Email Content</h4>
                                <div class="email-content">
                                    {5}
                                </div>
                            </div>
                        </div>
                    </div>
                    """.format(
                        message_id,
                        people_id,
                        getattr(a, 'Subject', 'No Subject'),
                        getattr(a, 'FromName', 'Unknown'),
                        getattr(a, 'Sent', ''),
                        getattr(a, 'Body', 'No content')
                    )
                
                # If no data found
                return "<div class='alert alert-warning'>Email not found</div>"
                
            except Exception as e:
                return "<div class='alert alert-danger'>Error loading email preview: {0}</div>".format(str(e))

        def render_transaction_history(self, people_id):
            """Render transaction history for a person"""
            transactions = self.get_transaction_history(people_id)
            
            # Get person name
            person_name = "Unknown Person"
            try:
                person_data = q.QuerySql("SELECT Name FROM People WHERE PeopleId = {0}".format(people_id))
                if person_data and len(person_data) > 0:
                    person_name = self.safe_get_attr(person_data[0], 'Name', 'Unknown Person')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-history"></i> Transaction History - {0}</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="history.go(-1)">
                            <i class="fa fa-arrow-left"></i> Go Back
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-search">
                    <input type="text" id="transactionSearch" placeholder="Search transactions..." 
                           class="form-control">
                </div>
                
                <div class="pm-content">
                    <table class="pm-table" id="transactionsTable">
                        <thead>
                            <tr>
                                <th>Date</th>
                                <th>Organization</th>
                                <th>Amount</th>
                                <th>Type</th>
                                <th>Direction</th>
                                <th>Description</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(person_name)
            
            try:
                if not transactions:
                    html += "<tr><td colspan='6' class='pm-center'>No transaction history found</td></tr>"
                else:
                    for trans in transactions:
                        trans_date = self.safe_get_attr(trans, 'TransactionDate', '')
                        org_name = self.safe_get_attr(trans, 'OrgName', 'Unknown')
                        amount = float(self.safe_get_attr(trans, 'Amount', 0) or 0)
                        trans_type = self.safe_get_attr(trans, 'TransactionType', 'Unknown')
                        direction = self.safe_get_attr(trans, 'TransactionDirection', '')
                        description = self.safe_get_attr(trans, 'Description', '')
                        message = self.safe_get_attr(trans, 'Message', '')
                        
                        # Combine description and message
                        full_desc = description
                        if message and message != description:
                            full_desc = description + " (" + message + ")" if description else message
                        
                        # Color code amounts
                        amount_class = "pm-positive" if amount > 0 else "pm-negative"
                        amount_display = self.format_currency(abs(amount))
                        
                        html += """
                                    <tr>
                                        <td>{0}</td>
                                        <td>{1}</td>
                                        <td class="pm-currency {2}">{3}</td>
                                        <td>{4}</td>
                                        <td>{5}</td>
                                        <td>{6}</td>
                                    </tr>
                        """.format(
                            trans_date,
                            org_name,
                            amount_class,
                            amount_display,
                            trans_type,
                            direction,
                            full_desc
                        )
            except Exception as e:
                html += "<tr><td colspan='6'>Error loading transaction history: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                    </table>
                </div>
            </div>
            """
            
            return html

        # ==============================================================
        # RECEIPTS VIEW -- date-range search of past check/cash payments
        # with per-row Reprint (print popup) + Email (resend) buttons.
        # All work happens client-side via the find_receipts and
        # reprint_receipt AJAX actions.
        # ==============================================================
        def render_receipts_view(self):
            today = datetime.now().strftime('%Y-%m-%d')
            # Default to a 7-day window so the typical "what did I record
            # this week?" lookup needs zero typing.
            try:
                from datetime import timedelta
                seven_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
            except:
                seven_ago = today
            return """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-receipt"></i> Reprint Receipts</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="viewPrograms()">
                            <i class="fa fa-arrow-left"></i> Back
                        </button>
                    </div>
                </div>
                <div class="pm-content">
                    <p style="color:#666;font-size:13px;margin:0 0 14px;">
                        Find past payments by date &mdash; <strong>check</strong>, <strong>cash</strong>,
                        <strong>credit card</strong>, <strong>coupon</strong>, or <strong>other</strong>
                        (manual entries, ACH, wire transfers, anything else that reduced a
                        balance). Reprint a popup receipt for the church files, re-email it to
                        the payer, or <strong>Customize</strong> to add a personal note and send
                        to any email. Re-emailed receipts use the original template with a
                        "Duplicate Receipt" banner so the recipient knows it isn't a new charge.
                        <em>Email is only enabled for check / cash payments recorded through
                        Payment Manager</em>; all others are view / print or Customize-and-send.
                    </p>
                    <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:end;
                                background:#f8f9fa;border:1px solid #e0e0e0;
                                border-radius:8px;padding:12px 14px;margin-bottom:14px;">
                        <div>
                            <label style="display:block;font-size:12px;color:#666;
                                          font-weight:600;margin-bottom:3px;">From date</label>
                            <input type="date" id="rcptFrom" value=\"""" + seven_ago + """\"
                                   style="padding:6px 8px;border:1px solid #ccc;border-radius:4px;">
                        </div>
                        <div>
                            <label style="display:block;font-size:12px;color:#666;
                                          font-weight:600;margin-bottom:3px;">To date</label>
                            <input type="date" id="rcptTo" value=\"""" + today + """\"
                                   style="padding:6px 8px;border:1px solid #ccc;border-radius:4px;">
                        </div>
                        <div>
                            <label style="display:block;font-size:12px;color:#666;
                                          font-weight:600;margin-bottom:3px;">Payment type</label>
                            <div id="rcptTypeDD" style="position:relative;display:inline-block;">
                                <button type="button" id="rcptTypeBtn" onclick="pmRcptTypeToggle(event)" style="cursor:pointer;padding:6px 28px 6px 10px;border:1px solid #ccc;border-radius:4px;background:#fff;min-width:200px;font-size:13px;position:relative;text-align:left;">
                                    <span id="rcptTypeLabel">All types</span>
                                    <i class="fa fa-caret-down" style="position:absolute;right:10px;top:50%;transform:translateY(-50%);color:#666;"></i>
                                </button>
                                <div id="rcptTypePanel" style="display:none;position:fixed;top:0;left:0;background:#fff;border:1px solid #ccc;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,0.16);padding:8px 10px;z-index:99999;min-width:200px;">
                                    <label style="display:flex;align-items:center;gap:6px;padding:3px 0;cursor:pointer;font-size:13px;font-weight:400;">
                                        <input type="checkbox" class="rcpt-type-cb" value="check" checked> Check
                                    </label>
                                    <label style="display:flex;align-items:center;gap:6px;padding:3px 0;cursor:pointer;font-size:13px;font-weight:400;">
                                        <input type="checkbox" class="rcpt-type-cb" value="cash" checked> Cash
                                    </label>
                                    <label style="display:flex;align-items:center;gap:6px;padding:3px 0;cursor:pointer;font-size:13px;font-weight:400;">
                                        <input type="checkbox" class="rcpt-type-cb" value="credit" checked> Credit Card
                                    </label>
                                    <label style="display:flex;align-items:center;gap:6px;padding:3px 0;cursor:pointer;font-size:13px;font-weight:400;">
                                        <input type="checkbox" class="rcpt-type-cb" value="coupon" checked> Coupon
                                    </label>
                                    <label style="display:flex;align-items:center;gap:6px;padding:3px 0;cursor:pointer;font-size:13px;font-weight:400;" title="Payments TouchPoint didn't prefix with CHK|/CSH| (manual entries, ACH, wire transfers, etc.)">
                                        <input type="checkbox" class="rcpt-type-cb" value="other" checked> Other
                                    </label>
                                    <div style="border-top:1px solid #eee;margin:4px -2px 2px;"></div>
                                    <label style="display:flex;align-items:center;gap:6px;padding:3px 0;cursor:pointer;font-size:13px;font-weight:400;color:#a8071a;" title="Money-out reversals (refunds). Check this with another type to see both payments AND refunds; check alone to audit refunds only; uncheck to hide all refunds.">
                                        <input type="checkbox" class="rcpt-type-cb" value="refund"> Refund
                                    </label>
                                </div>
                            </div>
                        </div>
                        <div>
                            <label style="display:block;font-size:12px;color:#666;font-weight:600;margin-bottom:3px;">Payer (optional)</label>
                            <input type="text" id="rcptPayer" placeholder="Name contains..." style="padding:6px 8px;border:1px solid #ccc;border-radius:4px;width:160px;" onkeydown="if(event.key==='Enter') findReceipts();">
                        </div>
                        <div>
                            <label style="display:block;font-size:12px;color:#666;font-weight:600;margin-bottom:3px;">Involvement (optional)</label>
                            <input type="text" id="rcptOrg" placeholder="Org name contains..." style="padding:6px 8px;border:1px solid #ccc;border-radius:4px;width:180px;" onkeydown="if(event.key==='Enter') findReceipts();">
                        </div>
                        <div>
                            <label style="display:block;font-size:12px;color:#666;font-weight:600;margin-bottom:3px;">Last 4 (CC/ACH)</label>
                            <input type="text" id="rcptLast4" placeholder="1234" inputmode="numeric" maxlength="4" style="padding:6px 8px;border:1px solid #ccc;border-radius:4px;width:90px;font-family:Menlo,Consolas,monospace;" onkeydown="if(event.key==='Enter') findReceipts();" title="Search by the last 4 digits of the card or bank account number stored on the transaction">
                        </div>
                        <button class="btn btn-primary" onclick="findReceipts()">
                            <i class="fa fa-search"></i> Find Receipts
                        </button>
                        <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:#666;cursor:pointer;margin-left:8px;">
                            <input type="checkbox" id="rcptIncludeExternal" checked>
                            <span title="Include payments recorded outside this tool (TouchPoint batch entry, syncs, etc.). Email reprint is disabled for those since no original receipt was ever sent.">
                                Include payments not recorded via Payment Manager
                            </span>
                        </label>
                    </div>
                    <div id="rcptResults" style="font-size:13px;color:#666;">
                        Pick a date range and click <strong>Find Receipts</strong>.
                    </div>
                </div>
            </div>
            """

        def _build_setup_banner_html(self):
            """Yellow callout listing missing/incomplete settings. Renders
            empty when everything is set up so it doesn't pollute the page."""
            try:
                issues = pm_setup_issues()
            except:
                return ''
            if not issues:
                return ''
            rows = []
            for problem, consequence in issues:
                rows.append(
                    '<li style="margin-bottom:4px;"><strong>' + problem + '</strong> '
                    '<span style="color:#7a5c00;font-style:italic;">' + consequence + '</span></li>'
                )
            return (
                '<div id="pmConfigBanner" style="max-width:1200px;margin:8px auto 0;padding:0 14px;">'
                '<div style="background:#fff4ce;border:1px solid #f4d35e;border-radius:8px;padding:12px 16px;color:#5a4500;font-size:13px;">'
                '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:14px;">'
                '<div style="flex:1;">'
                '<div style="font-weight:700;margin-bottom:6px;font-size:14px;">'
                '&#9888;&#65039; Payment Manager setup incomplete</div>'
                '<ul style="margin:0 0 6px 22px;padding:0;font-size:13px;">' + ''.join(rows) + '</ul>'
                '<div style="font-size:11px;color:#7a5c00;font-style:italic;">'
                'Click <strong>Settings</strong> above to fix. The drill-down still works, but the items above are limited until configured.</div>'
                '</div>'
                '<button onclick="openSettings()" style="background:#1f4e79;color:#fff;border:0;padding:6px 14px;border-radius:4px;cursor:pointer;font-weight:600;white-space:nowrap;">&#9881; Open Settings</button>'
                '</div></div></div>'
            )

        def render_page_structure(self, content):
            """Render the complete page with navigation and styling"""
            return """
            <script>
            // Helper function to get the correct form submission URL
            function getPyScriptAddress() {{
                let path = window.location.pathname;
                return path.replace("/PyScript/", "/PyScriptForm/");
            }}
            
            // Define all functions globally - before any HTML that uses them
            function showLoading() {{
                var loading = document.getElementById('pmLoading');
                if (loading) loading.style.display = 'flex';
            }}
            
            function hideLoading() {{
                var loading = document.getElementById('pmLoading');
                if (loading) loading.style.display = 'none';
            }}
            
            function showAlert(message, type) {{
                type = type || 'success';
                var alertContainer = document.getElementById('alertContainer');
                if (!alertContainer) return;
                
                var alertDiv = document.createElement('div');
                alertDiv.className = 'alert alert-' + type;
                alertDiv.innerHTML = '<button type="button" style="float: right; background: none; border: none; font-size: 18px; cursor: pointer;" onclick="this.parentNode.remove()">&times;</button>' + message;
                alertContainer.appendChild(alertDiv);
                
                setTimeout(function() {{
                    if (alertDiv.parentNode) {{
                        alertDiv.parentNode.removeChild(alertDiv);
                    }}
                }}, 5000);
            }}
            
            function viewPrograms() {{
                showLoading();
                window.location.href = window.location.pathname + '?action=programs';
            }}
            
            function viewDivisions(programId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=divisions&ProgramID=' + programId;
            }}
            
            function viewPayers(orgId, programId) {{
                showLoading();
                var url = window.location.pathname + '?action=payers';
                if (orgId && orgId !== 'null') url += '&OrganizationId=' + orgId;
                if (programId && programId !== 'null') url += '&ProgramID=' + programId;
                window.location.href = url;
            }}
            
            // --- Modernization additions: Receipts, Settings, Help, Auto-update -----
            function viewReceipts() {{
                showLoading();
                window.location.href = window.location.pathname + '?action=receipts';
            }}

            function pmEsc(s) {{
                if (s === null || s === undefined) return '';
                return String(s)
                    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
            }}

            function pmFmtMoney(n) {{
                var v = Math.abs(parseFloat(n) || 0);
                return '$' + v.toFixed(2).replace(/\\B(?=(\\d{{3}})+(?!\\d))/g, ',');
            }}

            // Toggle the Payment Type dropdown panel. Uses position:fixed
            // so the panel escapes ALL parent overflow / clipping
            // contexts; we compute its viewport coords from the button's
            // bounding rect each time it opens.
            function pmRcptTypeToggle(evt) {{
                if (evt) {{ evt.preventDefault(); evt.stopPropagation(); }}
                var panel = document.getElementById('rcptTypePanel');
                var btn   = document.getElementById('rcptTypeBtn');
                if (!panel || !btn) return;
                if (panel.style.display === 'block') {{
                    panel.style.display = 'none';
                    return;
                }}
                var rect = btn.getBoundingClientRect();
                panel.style.top   = (rect.bottom + 4) + 'px';
                panel.style.left  = rect.left + 'px';
                panel.style.minWidth = rect.width + 'px';
                panel.style.display = 'block';
            }}
            // Re-anchor on scroll/resize while open so the panel
            // tracks the button.
            window.addEventListener('scroll', function() {{
                var panel = document.getElementById('rcptTypePanel');
                var btn   = document.getElementById('rcptTypeBtn');
                if (!panel || !btn || panel.style.display !== 'block') return;
                var rect = btn.getBoundingClientRect();
                panel.style.top  = (rect.bottom + 4) + 'px';
                panel.style.left = rect.left + 'px';
            }}, true);
            // Close on outside click. Listens at document level; we
            // don't fire when the click landed inside #rcptTypeDD so
            // checkbox toggles still work.
            document.addEventListener('click', function(e) {{
                var dd = document.getElementById('rcptTypeDD');
                var panel = document.getElementById('rcptTypePanel');
                if (!dd || !panel) return;
                if (panel.style.display !== 'block') return;
                if (!dd.contains(e.target)) panel.style.display = 'none';
            }});
            // Close on Esc.
            document.addEventListener('keydown', function(e) {{
                if (e.key === 'Escape') {{
                    var panel = document.getElementById('rcptTypePanel');
                    if (panel && panel.style.display === 'block') panel.style.display = 'none';
                }}
            }});

            function pmRcptCollectTypes() {{
                // Read multi-select dropdown -> comma-separated codes.
                // Also keeps the summary label readable ("Check, Cash" /
                // "All types" / "2 selected").
                var boxes = document.querySelectorAll('.rcpt-type-cb');
                var picked = [];
                var labels = [];
                var nameByVal = {{ check:'Check', cash:'Cash', credit:'Credit Card', coupon:'Coupon', other:'Other', refund:'Refund' }};
                for (var i = 0; i < boxes.length; i++) {{
                    if (boxes[i].checked) {{
                        picked.push(boxes[i].value);
                        labels.push(nameByVal[boxes[i].value] || boxes[i].value);
                    }}
                }}
                var summary = document.getElementById('rcptTypeLabel');
                if (summary) {{
                    if (picked.length === 0) summary.textContent = 'None selected';
                    else if (picked.length === boxes.length) summary.textContent = 'All types';
                    else if (picked.length <= 2) summary.textContent = labels.join(', ');
                    else summary.textContent = picked.length + ' selected';
                }}
                return picked.join(',');
            }}

            // Live-update the summary label whenever a checkbox flips.
            document.addEventListener('change', function(e) {{
                if (e.target && e.target.classList && e.target.classList.contains('rcpt-type-cb')) {{
                    pmRcptCollectTypes();
                }}
            }});

            function findReceipts() {{
                var dFrom = document.getElementById('rcptFrom').value;
                var dTo   = document.getElementById('rcptTo').value;
                var rTypes = pmRcptCollectTypes();
                var incExt = document.getElementById('rcptIncludeExternal');
                var includeExternal = incExt && incExt.checked;
                var out = document.getElementById('rcptResults');
                if (!dFrom || !dTo) {{ showAlert('Pick both From and To dates', 'warning'); return; }}
                if (!rTypes) {{ showAlert('Pick at least one payment type', 'warning'); return; }}
                out.innerHTML = '<div style="padding:12px;color:#666;"><i class="fa fa-spinner fa-spin"></i> Searching...</div>';
                var payerEl = document.getElementById('rcptPayer');
                var orgEl   = document.getElementById('rcptOrg');
                var last4El = document.getElementById('rcptLast4');
                var fd = new FormData();
                fd.append('action', 'find_receipts');
                fd.append('date_from', dFrom);
                fd.append('date_to', dTo);
                fd.append('rcpt_types', rTypes);
                fd.append('include_external', includeExternal ? 'true' : 'false');
                if (payerEl && payerEl.value.trim()) fd.append('payer', payerEl.value.trim());
                if (orgEl && orgEl.value.trim())     fd.append('involvement', orgEl.value.trim());
                if (last4El && last4El.value.trim()) fd.append('last4', last4El.value.trim());
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r) {{ return r.text(); }})
                    .then(function(txt) {{
                        var d; try {{ d = JSON.parse(txt); }} catch(e) {{ throw new Error('Bad JSON'); }}
                        if (!d.success) {{ out.innerHTML = '<div class="alert alert-danger">' + pmEsc(d.message) + '</div>'; return; }}
                        var rows = (d.receipts || []);
                        if (!rows.length) {{
                            // Diagnostic empty-state -- explain WHY no rows
                            // came back so the staffer doesn't have to
                            // guess between "no data" and "filter bug".
                            var diag = d.diag || {{}};
                            var total = diag.totalTxns || 0;
                            var nonZero = diag.nonZeroAmt || 0;
                            var chk = diag.chkCount || 0;
                            var csh = diag.cshCount || 0;
                            var cc  = diag.ccCount || 0;
                            var cpn = diag.couponCount || 0;
                            var pm  = diag.pmCount || 0;
                            var ext = diag.extCount || 0;
                            var latest = diag.latestTxnDate || '';
                            var tableTotal = diag.tableTotal || 0;
                            var probeErr = diag.tableProbeErr || '';
                            var dfEcho = diag.dateFromEcho || '';
                            var dtEcho = diag.dateToEcho || '';
                            var inc = !!d.includeExternal;
                            var msg = '<div style="padding:20px;text-align:center;color:#666;">';
                            msg += '<div style="font-weight:600;margin-bottom:8px;">No matching receipts found in this date range.</div>';
                            msg += '<div style="font-size:12px;line-height:1.6;text-align:left;display:inline-block;background:#f8f9fa;border:1px solid #e1e5eb;border-radius:6px;padding:10px 14px;">';
                            msg += '<b>Diagnostic</b><br/>';
                            msg += 'Date range searched: <b>' + pmEsc(dfEcho) + '</b> to <b>' + pmEsc(dtEcho) + '</b><br/>';
                            if (diag.rawDateFrom !== undefined) {{
                                msg += '<span style="color:#999;">Raw POST values: from=' + pmEsc(diag.rawDateFrom) + ' to=' + pmEsc(diag.rawDateTo) + '</span><br/>';
                            }}
                            msg += 'Transactions in range (any amount): <b>' + total + '</b><br/>';
                            msg += 'With a non-zero amount: <b>' + nonZero + '</b><br/>';
                            var oth = diag.otherCount || 0;
                            msg += 'Check: <b>' + chk + '</b> &middot; Cash: <b>' + csh + '</b> &middot; Credit Card: <b>' + cc + '</b> &middot; Coupon: <b>' + cpn + '</b> &middot; Other: <b>' + oth + '</b><br/>';
                            msg += 'Of CHK/CSH, Payment Manager: <b>' + pm + '</b> &middot; External: <b>' + ext + '</b><br/>';
                            msg += 'Whole table: <b>' + tableTotal + '</b> rows';
                            if (latest) msg += ' &middot; latest <b>' + pmEsc(latest) + '</b>';
                            msg += '<br/>';
                            // Decision tree -- ordered most specific to most generic.
                            if (probeErr) {{
                                msg += '<div style="margin-top:8px;color:#a94442;">Database probe failed: <code>' + pmEsc(probeErr) + '</code>. The script may not have permission to read [Transaction] in this environment.</div>';
                            }} else if (tableTotal === 0) {{
                                msg += '<div style="margin-top:8px;color:#a94442;">The [Transaction] table appears empty (0 total rows). You are likely connected to a sandbox or test database.</div>';
                            }} else if (total === 0) {{
                                msg += '<div style="margin-top:8px;color:#7a5c00;">No transactions exist in the searched date range. ';
                                if (latest) {{
                                    msg += 'The most recent transaction in the database is <b>' + pmEsc(latest) + '</b>';
                                    msg += ' &mdash; if that is earlier than your "From date", widen the range or this DB snapshot is stale.';
                                }} else {{
                                    msg += 'Widen the date range to verify.';
                                }}
                                msg += '</div>';
                            }} else if (nonZero === 0) {{
                                msg += '<div style="margin-top:8px;color:#7a5c00;">' + total + ' transactions exist but none have a non-zero amount. Receipts are only shown for payments with money attached.</div>';
                            }} else if (!inc && ext > 0) {{
                                msg += '<div style="margin-top:8px;color:#7a5c00;">Tip: ' + ext + ' external CHK/CSH payment(s) exist -- enable <i>Include payments not recorded via Payment Manager</i> to see them.</div>';
                            }} else if (chk + csh + cc + cpn + oth === 0) {{
                                msg += '<div style="margin-top:8px;">Transactions exist in this range, but none match any of the payment types.</div>';
                            }} else {{
                                msg += '<div style="margin-top:8px;">Transactions of the requested types exist but no rows came back. Check your payment-type selection.</div>';
                            }}
                            msg += '</div></div>';
                            out.innerHTML = msg;
                            return;
                        }}
                        var pmCount = 0, extCount = 0;
                        for (var c = 0; c < rows.length; c++) {{ if (rows[c].source === 'pm') pmCount++; else extCount++; }}
                        var html = '<div style="font-size:12px;color:#666;margin-bottom:6px;">'
                                 + rows.length + ' receipt(s) found &middot; '
                                 + pmCount + ' from Payment Manager';
                        if (extCount) html += ', ' + extCount + ' from elsewhere (Email disabled)';
                        html += '.</div>';
                        html += '<table class="pm-table"><thead><tr>';
                        html += '<th>Date</th><th>Type</th><th>Source</th><th>Payer</th><th style="text-align:right;">Amount</th><th>Involvement</th><th>Email on file</th><th style="white-space:nowrap;">Actions</th>';
                        html += '</tr></thead><tbody>';
                        for (var i = 0; i < rows.length; i++) {{
                            var r = rows[i];
                            var hasEmail = !!r.email;
                            var isPM = (r.source === 'pm');
                            html += '<tr' + (isPM ? '' : ' style="background:#fffaf0;"') + '>';
                            html += '<td>' + pmEsc(r.tranDate) + '</td>';
                            var ptColors = {{ 'Check':'#cfe3ff', 'Cash':'#d4edda', 'Credit Card':'#ffe6cc', 'Coupon':'#f3e5f5', 'Other Payment':'#e2e3e5',
                                              'Refund':'#f8d7da', 'Check Refund':'#f8d7da', 'Cash Refund':'#f8d7da' }};
                            var ptFgColors = {{ 'Refund':'#a8071a', 'Check Refund':'#a8071a', 'Cash Refund':'#a8071a' }};
                            var ptBg = ptColors[r.paymentType] || '#e9ecef';
                            var ptFg = ptFgColors[r.paymentType] || '#1f4e79';
                            var last4 = (r.last4cc || r.last4ach || '').toString().trim();
                            var last4Chip = '';
                            if (last4) {{
                                var l4Label = r.last4ach ? 'ACH ****' : '****';
                                last4Chip = '<div style="margin-top:2px;font-family:Menlo,Consolas,monospace;font-size:10px;color:#666;">' + l4Label + ' ' + pmEsc(last4) + '</div>';
                            }}
                            html += '<td><span style="background:' + ptBg + ';color:' + ptFg + ';padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;">' + pmEsc(r.paymentType) + '</span>' + last4Chip + '</td>';
                            // Source column: PM = blue badge, External =
                            // amber badge with tooltip explaining the
                            // Email button limitation.
                            if (isPM) {{
                                html += '<td><span style="background:#deecf9;color:#1f4e79;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">Payment Manager</span></td>';
                            }} else {{
                                html += '<td><span title="Recorded outside Payment Manager (no original receipt was sent). Email is disabled; you can still Print for the church files." style="background:#fff4ce;color:#7a5c00;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;cursor:help;">External</span></td>';
                            }}
                            // Payer cell -- name on top, then a sub-line
                            // that carries the most useful detail per
                            // type:
                            //   * Coupon  -> bare coupon code (TxnId)
                            //   * Other types -> staff note from Message,
                            //                    falls back to TxnId if
                            //                    truly nothing else
                            // Tooltip always carries the raw TransactionId
                            // for forensics.
                            var txIdShort = r.transactionId || '';
                            var couponMatch = /^Coupon\\((.+)\\)$/.exec(txIdShort);
                            var subLine = '';
                            var subStyle = 'font-size:10px;color:#999;font-family:Menlo,Consolas,monospace;margin-top:1px;';
                            if (couponMatch) {{
                                subLine = couponMatch[1];   // bare coupon code
                            }} else if ((r.note || '').trim()) {{
                                // Truncate long notes so the row stays tidy.
                                var n = r.note.trim();
                                subLine = n.length > 80 ? n.substring(0, 78) + '...' : n;
                                // Regular font for human-readable notes (not mono).
                                subStyle = 'font-size:11px;color:#777;font-style:italic;margin-top:1px;';
                            }} else {{
                                subLine = txIdShort;
                            }}
                            html += '<td>' + pmEsc(r.person) + '<div style="' + subStyle + '" title="Transaction ID: ' + pmEsc(r.transactionId) + (r.note ? ' | Full note: ' + pmEsc(r.note) : '') + '">' + pmEsc(subLine) + '</div></td>';
                            html += '<td style="text-align:right;font-weight:600;">' + pmFmtMoney(r.amount) + '</td>';
                            // Involvement cell -- always surface who was
                            // registered. Single name when it's just the
                            // payer's own registration; comma-separated
                            // list when a family registered together.
                            var orgCell = pmEsc(r.orgName);
                            var regs = (r.registrants || '').trim();
                            if (regs) {{
                                var icon = (regs.indexOf(',') >= 0) ? 'fa-users' : 'fa-user';
                                orgCell += '<div style="font-size:11px;color:#777;margin-top:2px;"><i class="fa ' + icon + '" style="opacity:0.6;"></i> Registered: ' + pmEsc(regs) + '</div>';
                            }}
                            html += '<td style="font-size:12px;color:#666;">' + orgCell + '</td>';
                            html += '<td style="font-size:12px;">' + (hasEmail ? pmEsc(r.email) : '<span style="color:#999;font-style:italic;">none</span>') + '</td>';
                            html += '<td style="white-space:nowrap;">';
                            var escTxId = pmEsc(r.transactionId).replace(/\\\\/g, '\\\\\\\\').replace(/\\'/g, "\\\\\\'");
                            html += '<button class="btn btn-sm btn-secondary" onclick="reprintReceiptPrint(\\'' + escTxId + '\\')" style="padding:3px 8px;font-size:11px;"><i class="fa fa-print"></i> Print</button> ';
                            // Email + Customize buttons: both require
                            // PM-origin AND a valid email-on-file.
                            var emailDisabledReason = '';
                            var isRefund = (String(r.paymentType || '').indexOf('Refund') >= 0);
                            if (isRefund) emailDisabledReason = 'This is a refund -- emailing a payment receipt would be misleading';
                            else if (!isPM) emailDisabledReason = 'Recorded outside Payment Manager -- no original receipt was sent';
                            else if (!hasEmail) emailDisabledReason = 'No email on file';
                            var emailDisabled = !!emailDisabledReason;
                            html += '<button class="btn btn-sm btn-primary" onclick="reprintReceiptEmail(\\'' + escTxId + '\\')" style="padding:3px 8px;font-size:11px;"' + (emailDisabled ? ' disabled title="' + pmEsc(emailDisabledReason) + '"' : '') + '><i class="fa fa-envelope"></i> Email</button> ';
                            // Customize is always enabled -- it generates
                            // a fresh receipt the staffer edits and then
                            // prints or sends to any email. Unlike the
                            // direct Email button, it's not a duplicate
                            // of an original PM send.
                            html += '<button class="btn btn-sm" onclick="reprintReceiptCustomize(\\'' + escTxId + '\\')" style="padding:3px 8px;font-size:11px;background:#6c757d;color:#fff;border:none;" title="Edit the receipt before sending or printing -- add a personal note, change wording, send to any email."><i class="fa fa-edit"></i> Customize</button>';
                            html += '</td>';
                            html += '</tr>';
                        }}
                        html += '</tbody></table>';
                        out.innerHTML = html;
                    }})
                    .catch(function(err) {{ out.innerHTML = '<div class="alert alert-danger">Search failed: ' + pmEsc(err.message) + '</div>'; }});
            }}

            function _reprintInvoke(txId, mode, onOk) {{
                var fd = new FormData();
                fd.append('action', 'reprint_receipt');
                fd.append('transaction_id', txId);
                fd.append('mode', mode);
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r) {{ return r.text(); }})
                    .then(function(txt) {{
                        var d; try {{ d = JSON.parse(txt); }} catch(e) {{ showAlert('Bad response from server', 'danger'); return; }}
                        if (!d.success) {{ showAlert(d.message || 'Failed', 'danger'); return; }}
                        if (onOk) onOk(d);
                    }})
                    .catch(function(err) {{ showAlert('Network error: ' + err.message, 'danger'); }});
            }}

            function reprintReceiptEmail(txId) {{
                if (!confirm('Re-email this receipt to the payer on file?')) return;
                _reprintInvoke(txId, 'email', function(d) {{ showAlert(d.message || 'Sent', 'success'); }});
            }}

            function reprintReceiptPrint(txId) {{
                _reprintInvoke(txId, 'html', function(d) {{
                    var body = (d.html) || '';
                    var w = window.open('', '_blank');
                    if (!w) {{ showAlert('Popup blocked -- allow popups to print', 'warning'); return; }}
                    w.document.write('<!DOCTYPE html><html><head><title>' + pmEsc(d.subject || 'Receipt') + '</title>');
                    w.document.write('<style>*{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}}body{{font-family:Segoe UI,Tahoma,sans-serif;color:#333;max-width:680px;margin:24px auto;padding:0 18px;}}</style>');
                    w.document.write('</head><body>' + body + '</body></html>');
                    w.document.close();
                    w.focus();
                    setTimeout(function(){{ w.print(); }}, 300);
                }});
            }}

            // --- Customize-before-send: load the rendered receipt into
            // an editable modal so the staffer can add language
            // ("thank you for your gift", etc.) before firing it.
            var pmCustomTxId = '';
            var pmCustomOnFileEmail = '';  // email-on-file, used by "Use email on file"
            var pmCustomReceiptBody = ''; // immutable receipt HTML loaded once; we never edit it,
                                          // we just prepend the personal note on send + preview
            function reprintReceiptCustomize(txId) {{
                pmCustomTxId = txId;
                pmCustomOnFileEmail = '';
                pmCustomReceiptBody = '';
                document.getElementById('pmCustomTo').textContent      = 'Loading...';
                document.getElementById('pmCustomSubject').value       = '';
                document.getElementById('pmCustomNote').value          = '';
                document.getElementById('pmCustomToEmail').value       = '';
                document.getElementById('pmCustomToHint').textContent  = '(loading...)';
                document.getElementById('pmCustomPreview').innerHTML   = '<div style="padding:18px;color:#999;">Loading receipt...</div>';
                document.getElementById('pmCustomModal').style.display = 'flex';
                _reprintInvoke(txId, 'html', function(d) {{
                    pmCustomOnFileEmail = d.email || '';
                    pmCustomReceiptBody = d.html || '';
                    document.getElementById('pmCustomTo').textContent    = d.person || '';
                    document.getElementById('pmCustomSubject').value     = d.subject || 'Receipt (duplicate)';
                    document.getElementById('pmCustomToEmail').value     = pmCustomOnFileEmail;
                    pmCustomUpdateEmailHint();
                    pmCustomRefreshPreview();
                }});
            }}
            function pmCustomNoteToHtml(noteText) {{
                // Plain-text note -> HTML: blank lines become <p>, single
                // newlines become <br>. Escape HTML so the note never
                // injects markup.
                if (!noteText) return '';
                var esc = pmEsc(noteText);
                var paragraphs = esc.split(/\\n\\s*\\n/);
                var html = paragraphs.map(function(p) {{
                    return '<p style="margin:0 0 10px;">' + p.replace(/\\n/g, '<br>') + '</p>';
                }}).join('');
                // Wrap in a friendly callout so the note stands out
                // visually from the rest of the receipt.
                return '<div style="background:#eef5fb;border-left:4px solid #1f4e79;padding:12px 16px;margin:0 0 16px;border-radius:0 6px 6px 0;font-size:14px;color:#1f4e79;line-height:1.5;">' + html + '</div>';
            }}
            function pmCustomUpdateEmailHint() {{
                // Show whether the current "Send to" matches the email
                // on file, or is an override.
                var typed = (document.getElementById('pmCustomToEmail').value || '').trim().toLowerCase();
                var onFile = (pmCustomOnFileEmail || '').toLowerCase();
                var hint = document.getElementById('pmCustomToHint');
                if (!hint) return;
                if (!typed) {{
                    hint.textContent = onFile ? '(defaults to ' + pmCustomOnFileEmail + ')' : '(no email on file)';
                    hint.style.color = '#999';
                }} else if (typed === onFile) {{
                    hint.textContent = '(using email on file)';
                    hint.style.color = '#999';
                }} else {{
                    hint.textContent = '(override -- not the email on file)';
                    hint.style.color = '#7a5c00';
                }}
            }}
            function pmCustomResetEmail() {{
                document.getElementById('pmCustomToEmail').value = pmCustomOnFileEmail;
                pmCustomUpdateEmailHint();
            }}
            document.addEventListener('input', function(e) {{
                if (e.target && e.target.id === 'pmCustomToEmail') pmCustomUpdateEmailHint();
            }});
            function pmCustomClose() {{
                document.getElementById('pmCustomModal').style.display = 'none';
                pmCustomTxId = '';
            }}
            function pmCustomComposedBody() {{
                // Note (if any) prepended above the receipt body. Used
                // by preview, print, and as the rendering reference for
                // the server (which composes the same way).
                var noteHtml = pmCustomNoteToHtml(document.getElementById('pmCustomNote').value || '');
                return noteHtml + (pmCustomReceiptBody || '');
            }}
            function pmCustomRefreshPreview() {{
                document.getElementById('pmCustomPreview').innerHTML = pmCustomComposedBody();
            }}
            function pmCustomPrint() {{
                var body = pmCustomComposedBody();
                var subject = document.getElementById('pmCustomSubject').value || 'Receipt';
                if (!body.trim()) {{ showAlert('Receipt is still loading', 'warning'); return; }}
                var w = window.open('', '_blank');
                if (!w) {{ showAlert('Popup blocked -- allow popups to print', 'warning'); return; }}
                w.document.write('<!DOCTYPE html><html><head><title>' + pmEsc(subject) + '</title>');
                w.document.write('<style>*{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}}body{{font-family:Segoe UI,Tahoma,sans-serif;color:#333;max-width:680px;margin:24px auto;padding:0 18px;}}</style>');
                w.document.write('</head><body>' + body + '</body></html>');
                w.document.close();
                w.focus();
                setTimeout(function(){{ w.print(); }}, 300);
            }}
            function pmCustomSend() {{
                if (!pmCustomTxId) {{ showAlert('No transaction context', 'danger'); return; }}
                var subject = document.getElementById('pmCustomSubject').value.trim();
                var body    = pmCustomComposedBody();
                var altEmail = document.getElementById('pmCustomToEmail').value.trim();
                if (!altEmail) {{ showAlert('Recipient email is required', 'warning'); return; }}
                // Basic email format check -- server validates again.
                if (!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(altEmail)) {{
                    showAlert('That doesn\\'t look like a valid email address', 'warning');
                    return;
                }}
                if (!subject) {{ showAlert('Subject is required', 'warning'); return; }}
                if (!body.trim()) {{ showAlert('Receipt is still loading', 'warning'); return; }}
                var isOverride = altEmail.toLowerCase() !== (pmCustomOnFileEmail || '').toLowerCase();
                var confirmMsg = 'Send this customized receipt to ' + altEmail + '?';
                if (isOverride && pmCustomOnFileEmail) {{
                    confirmMsg += '\\n\\nNote: this is NOT the email on file (' + pmCustomOnFileEmail + ').';
                }}
                if (!confirm(confirmMsg)) return;
                var btn = document.getElementById('pmCustomSendBtn');
                if (btn) {{ btn.disabled = true; btn.textContent = 'Sending...'; }}
                var fd = new FormData();
                fd.append('action', 'send_custom_receipt');
                fd.append('transaction_id', pmCustomTxId);
                fd.append('encoded', 'urlc');
                fd.append('subject', encodeURIComponent(subject));
                fd.append('body',    encodeURIComponent(body));
                fd.append('to_email', altEmail);
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{ return r.text(); }})
                    .then(function(txt){{
                        if (btn) {{ btn.disabled = false; btn.textContent = 'Send Customized Receipt'; }}
                        var d; try {{ d = JSON.parse(txt); }} catch(e) {{ showAlert('Bad response', 'danger'); return; }}
                        if (d.success) {{
                            showAlert(d.message || 'Sent', 'success');
                            pmCustomClose();
                        }} else {{
                            showAlert(d.message || 'Send failed', 'danger');
                        }}
                    }})
                    .catch(function(err){{
                        if (btn) {{ btn.disabled = false; btn.textContent = 'Send Customized Receipt'; }}
                        showAlert('Network error: ' + err.message, 'danger');
                    }});
            }}

            // --- Settings modal --------------------------------------------------
            var pmSettings = null;
            var pmTemplates = {{}};
            function openSettings() {{
                document.getElementById('pmSettingsModal').style.display = 'flex';
                var fd = new FormData();
                fd.append('action', 'load_settings');
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{ return r.text(); }})
                    .then(function(txt){{
                        try {{
                            var d = JSON.parse(txt);
                            if (d.success && d.settings) {{
                                pmSettings = d.settings;
                                pmTemplates = d.templates || {{}};
                                pmRenderSettings();
                            }} else {{
                                showAlert(d.message || 'Failed to load settings', 'danger');
                            }}
                        }} catch(e) {{ showAlert('Bad settings response', 'danger'); }}
                    }});
            }}
            function closeSettings() {{ document.getElementById('pmSettingsModal').style.display = 'none'; }}

            function pmRenderSettings() {{
                if (!pmSettings) return;
                var ds = pmSettings.defaultSender || {{}};
                document.getElementById('pmSetSenderId').value     = ds.sender_id || '';
                document.getElementById('pmSetSenderEmail').value  = ds.sender_email || '';
                document.getElementById('pmSetSenderAlias').value  = ds.sender_alias || '';
                document.getElementById('pmSetSenderTitle').value  = ds.email_title || '';
                document.getElementById('pmSetSenderPhone').value  = ds.sender_phone || '';
                document.getElementById('pmSetTplConfirm').value   = pmSettings.paymentConfirmationTemplate || '';
                document.getElementById('pmSetTplNotify').value    = pmSettings.paymentNotificationTemplate || '';
                document.getElementById('pmSetSubjectPrefix').value = pmSettings.emailSubjectPrefix || '';
                // Template editor pre-fill. pmTemplates was populated by
                // load_settings (server reads HtmlContent for each name
                // and falls back to defaults so the editor never opens
                // empty even before any custom template exists).
                var conf = pmTemplates.confirmation || {{}};
                var notif = pmTemplates.notification || {{}};
                var ce = document.getElementById('pmTplConfirmBody');
                var ne = document.getElementById('pmTplNotifyBody');
                if (ce) ce.value = conf.body || '';
                if (ne) ne.value = notif.body || '';
                var cn = document.getElementById('pmTplConfirmName');
                var nn = document.getElementById('pmTplNotifyName');
                if (cn) cn.textContent = conf.name || '';
                if (nn) nn.textContent = notif.name || '';
                var cb = document.getElementById('pmTplConfirmBadge');
                var nb = document.getElementById('pmTplNotifyBadge');
                if (cb) cb.textContent = conf.isDefault ? '(using built-in default)' : '(custom HtmlContent)';
                if (nb) nb.textContent = notif.isDefault ? '(using built-in default)' : '(custom HtmlContent)';
                pmRenderSendersTable();
            }}

            // --- Template editor save / reset -----------------------------------
            function pmSaveTemplate(which) {{
                var body = document.getElementById(which === 'confirmation' ? 'pmTplConfirmBody' : 'pmTplNotifyBody').value;
                var fd = new FormData();
                fd.append('action', 'save_template');
                fd.append('which', which);
                fd.append('encoded', 'urlc');
                fd.append('body', encodeURIComponent(body));
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{ return r.text(); }})
                    .then(function(txt){{
                        try {{
                            var d = JSON.parse(txt);
                            if (d.success) {{
                                showAlert(d.message || 'Saved', 'success');
                                // Refresh the "(custom HtmlContent)" badge
                                if (which === 'confirmation') {{
                                    pmTemplates.confirmation = pmTemplates.confirmation || {{}};
                                    pmTemplates.confirmation.isDefault = false;
                                    var cb = document.getElementById('pmTplConfirmBadge');
                                    if (cb) cb.textContent = '(custom HtmlContent)';
                                }} else {{
                                    pmTemplates.notification = pmTemplates.notification || {{}};
                                    pmTemplates.notification.isDefault = false;
                                    var nb = document.getElementById('pmTplNotifyBadge');
                                    if (nb) nb.textContent = '(custom HtmlContent)';
                                }}
                            }} else {{
                                showAlert(d.message || 'Save failed', 'danger');
                            }}
                        }} catch(e) {{ showAlert('Bad response', 'danger'); }}
                    }});
            }}

            function pmResetTemplate(which) {{
                if (!confirm('Restore the built-in default for this template? Your current changes will be discarded (you can still Save to keep the default as your custom template).')) return;
                var fd = new FormData();
                fd.append('action', 'reset_template');
                fd.append('which', which);
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{ return r.text(); }})
                    .then(function(txt){{
                        try {{
                            var d = JSON.parse(txt);
                            if (d.success && d.body !== undefined) {{
                                document.getElementById(which === 'confirmation' ? 'pmTplConfirmBody' : 'pmTplNotifyBody').value = d.body;
                                showAlert('Default loaded into editor -- click Save to apply', 'info');
                            }}
                        }} catch(e) {{}}
                    }});
            }}

            function pmPreviewTemplate(which) {{
                var body = document.getElementById(which === 'confirmation' ? 'pmTplConfirmBody' : 'pmTplNotifyBody').value;
                // Sample merge values so the admin can see what the email
                // will actually look like to a payer.
                var sample = body
                    .replace(/\\{{name\\}}/g, 'Sample Payer')
                    .replace(/\\{{chargeNotes\\}}/g, '$50.00 -- Check payment on 2026-05-21<br><span style="color:#666;font-size:13px;">For: <strong>Student Summer Camp 2026</strong></span><br><span style="color:#666;font-size:13px;">Registered: <strong>Sample Kid One, Sample Kid Two</strong></span>')
                    .replace(/\\{{orgName\\}}/g, 'Student Summer Camp 2026')
                    .replace(/\\{{registrants\\}}/g, 'Sample Kid One, Sample Kid Two')
                    .replace(/\\{{note\\}}/g, 'check #1234 -- spring camp deposit')
                    .replace(/\\{{previousDue\\}}/g, '$50.00 ........ Previous Balance')
                    .replace(/\\{{totalDue\\}}/g, '$0.00')
                    .replace(/\\{{newTotalDue\\}}/g, '$0.00')
                    .replace(/\\{{paylink\\}}/g, '#sample-paylink')
                    .replace(/\\{{sender_alias\\}}/g, 'Sample Sender')
                    .replace(/\\{{sender_phone\\}}/g, '(615) 555-1234')
                    .replace(/\\{{sender_email\\}}/g, 'sample@church.org');
                var w = window.open('', '_blank');
                if (!w) {{ showAlert('Popup blocked', 'warning'); return; }}
                w.document.write('<!DOCTYPE html><html><head><title>Template Preview</title>');
                w.document.write('<style>body{{font-family:Segoe UI,Tahoma,sans-serif;color:#333;max-width:680px;margin:24px auto;padding:0 18px;}}</style>');
                w.document.write('</head><body>');
                w.document.write('<div style="background:#eef5fb;border:1px solid #cfe3ff;padding:8px 12px;border-radius:6px;font-size:12px;color:#1f4e79;margin-bottom:14px;"><strong>Preview</strong> &mdash; merge fields filled with sample values. Real sends will use the actual payer name, amount, etc.</div>');
                w.document.write(sample);
                w.document.write('</body></html>');
                w.document.close();
            }}

            function pmRenderSendersTable() {{
                var tbody = document.getElementById('pmSendersTbody');
                var senders = (pmSettings && pmSettings.senders) || {{}};
                var html = '';
                var keys = Object.keys(senders);
                if (!keys.length) {{
                    html = '<tr><td colspan="6" style="padding:10px;color:#999;font-style:italic;text-align:center;">No per-program overrides. The Default sender above is used for every program.</td></tr>';
                }} else {{
                    keys.sort();
                    for (var i = 0; i < keys.length; i++) {{
                        var k = keys[i], s = senders[k] || {{}};
                        html += '<tr>';
                        html += '<td><input type="text" value="' + pmEsc(k) + '" data-key="' + pmEsc(k) + '" data-fld="_progId" style="width:70px;padding:3px 5px;"></td>';
                        html += '<td><input type="text" value="' + pmEsc(s.sender_id || '') + '" data-key="' + pmEsc(k) + '" data-fld="sender_id" style="width:70px;padding:3px 5px;"></td>';
                        html += '<td><input type="text" value="' + pmEsc(s.sender_email || '') + '" data-key="' + pmEsc(k) + '" data-fld="sender_email" style="width:100%;padding:3px 5px;"></td>';
                        html += '<td><input type="text" value="' + pmEsc(s.sender_alias || '') + '" data-key="' + pmEsc(k) + '" data-fld="sender_alias" style="width:100%;padding:3px 5px;"></td>';
                        html += '<td><input type="text" value="' + pmEsc(s.email_title || '') + '" data-key="' + pmEsc(k) + '" data-fld="email_title" style="width:100%;padding:3px 5px;"></td>';
                        html += '<td><button class="btn btn-sm" onclick="pmRemoveSender(\\'' + pmEsc(k) + '\\')" style="padding:2px 8px;color:#d13438;">Remove</button></td>';
                        html += '</tr>';
                    }}
                }}
                tbody.innerHTML = html;
            }}

            function pmAddSender() {{
                if (!pmSettings) return;
                var pid = prompt('Program ID for this sender override:');
                if (!pid) return;
                pmSettings.senders = pmSettings.senders || {{}};
                if (pmSettings.senders[pid]) {{ showAlert('Program ID already has an override', 'warning'); return; }}
                pmSettings.senders[pid] = {{ sender_id:'', sender_email:'', sender_alias:'', email_title:'', sender_phone:'' }};
                pmRenderSendersTable();
            }}

            function pmRemoveSender(key) {{
                if (!pmSettings || !pmSettings.senders) return;
                if (!confirm('Remove the sender override for Program ' + key + '?')) return;
                delete pmSettings.senders[key];
                pmRenderSendersTable();
            }}

            function pmSaveSettings() {{
                if (!pmSettings) {{ showAlert('Settings not loaded yet -- close and reopen the dialog', 'warning'); return; }}
                try {{
                    // Read all inputs back into pmSettings
                    pmSettings.defaultSender = pmSettings.defaultSender || {{}};
                    pmSettings.defaultSender.sender_id    = document.getElementById('pmSetSenderId').value;
                    pmSettings.defaultSender.sender_email = document.getElementById('pmSetSenderEmail').value;
                    pmSettings.defaultSender.sender_alias = document.getElementById('pmSetSenderAlias').value;
                    pmSettings.defaultSender.email_title  = document.getElementById('pmSetSenderTitle').value;
                    pmSettings.defaultSender.sender_phone = document.getElementById('pmSetSenderPhone').value;
                    pmSettings.paymentConfirmationTemplate  = document.getElementById('pmSetTplConfirm').value;
                    pmSettings.paymentNotificationTemplate = document.getElementById('pmSetTplNotify').value;
                    pmSettings.emailSubjectPrefix          = document.getElementById('pmSetSubjectPrefix').value;
                    // Read per-program rows (may have renamed program IDs)
                    var newSenders = {{}};
                    var inputs = document.querySelectorAll('#pmSendersTbody input');
                    var rowMap = {{}};
                    inputs.forEach(function(inp) {{
                        var k = inp.getAttribute('data-key'); var f = inp.getAttribute('data-fld');
                        if (!rowMap[k]) rowMap[k] = {{}};
                        rowMap[k][f] = inp.value;
                    }});
                    for (var k in rowMap) {{
                        var row = rowMap[k]; var newKey = (row._progId || k).toString();
                        delete row._progId;
                        if (newKey) newSenders[newKey] = row;
                    }}
                    pmSettings.senders = newSenders;
                }} catch(prepErr) {{
                    showAlert('Could not read settings form: ' + prepErr.message, 'danger');
                    return;
                }}
                var payload;
                try {{
                    payload = encodeURIComponent(JSON.stringify(pmSettings));
                }} catch(e) {{
                    showAlert('Could not serialize settings: ' + e.message, 'danger');
                    return;
                }}
                // URL-encode the JSON to dodge ASP.NET request validation
                // on any HTML the admin might put in a sender alias/title.
                var fd = new FormData();
                fd.append('action', 'save_settings');
                fd.append('encoded', 'urlc');
                fd.append('settings_json', payload);
                var btn = document.querySelector('button[onclick="pmSaveSettings()"]');
                if (btn) {{ btn.disabled = true; btn.textContent = 'Saving...'; }}
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{
                        if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + r.statusText);
                        return r.text();
                    }})
                    .then(function(txt){{
                        if (btn) {{ btn.disabled = false; btn.textContent = 'Save Settings'; }}
                        if (!txt) {{ showAlert('Server returned an empty response (often means ASP.NET request validation blocked the POST). Check the templates / sender fields for raw &lt;script&gt; tags.', 'danger'); return; }}
                        var d;
                        try {{ d = JSON.parse(txt); }}
                        catch(e) {{ showAlert('Server returned non-JSON (' + txt.substring(0, 120) + '...)', 'danger'); return; }}
                        if (d.success) {{
                            showAlert('Settings saved', 'success');
                            closeSettings();
                            // Reset the missing-settings banner if it was up.
                            var b = document.getElementById('pmConfigBanner');
                            if (b) b.style.display = 'none';
                        }} else {{
                            showAlert(d.message || 'Save failed (no message)', 'danger');
                        }}
                    }})
                    .catch(function(err){{
                        if (btn) {{ btn.disabled = false; btn.textContent = 'Save Settings'; }}
                        showAlert('Save network error: ' + err.message, 'danger');
                    }});
            }}

            // --- Help modal ------------------------------------------------------
            function openHelp() {{ document.getElementById('pmHelpModal').style.display = 'flex'; }}
            function closeHelp() {{ document.getElementById('pmHelpModal').style.display = 'none'; }}

            // --- Auto-update check ----------------------------------------------
            var PM_APP_VERSION = '{3}';
            var PM_DC_SCRIPT_ID = '{4}';
            var PM_DC_API_BASE = '{5}';
            function pmCheckUpdate() {{
                try {{
                    var xhr = new XMLHttpRequest();
                    xhr.open('GET', PM_DC_API_BASE + '/script-versions', true);
                    xhr.timeout = 5000;
                    xhr.onreadystatechange = function() {{
                        if (xhr.readyState !== 4 || xhr.status !== 200) return;
                        try {{
                            var v = JSON.parse(xhr.responseText)[PM_DC_SCRIPT_ID];
                            if (v && v !== PM_APP_VERSION) {{
                                var b = document.getElementById('pmUpdateBanner');
                                if (b) {{
                                    b.innerHTML = '<div style="background:#e8f0fd;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:12px;display:flex;align-items:center;gap:10px;">'
                                        + '<div style="font-size:18px;">&#128640;</div>'
                                        + '<div style="flex:1;font-size:12px;color:#0078d4;"><strong>Update available</strong> &mdash; you have <code>v' + PM_APP_VERSION + '</code>, latest is <code>v' + v + '</code>. Saved settings preserved.</div>'
                                        + '<button onclick="pmApplyUpdate()" style="padding:6px 14px;background:#0078d4;color:#fff;border:0;border-radius:4px;cursor:pointer;">Update Now</button>'
                                        + '</div>';
                                }}
                            }}
                        }} catch(e) {{}}
                    }};
                    xhr.send();
                }} catch(e) {{}}
            }}
            function pmApplyUpdate() {{
                if (!confirm('Update Payment Manager to the latest version?')) return;
                var fd = new FormData();
                fd.append('action', 'apply_update');
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{ return r.text(); }})
                    .then(function(txt){{
                        try {{
                            var d = JSON.parse(txt);
                            if (d.success) {{ alert(d.message || 'Updated'); window.location.reload(true); }}
                            else {{
                                // Use alert() not toast so the full hint
                                // (including manual-install URL) is
                                // visible. Toasts truncate.
                                alert(d.message || 'Update failed');
                            }}
                        }} catch(e) {{ showAlert('Bad update response', 'danger'); }}
                    }});
            }}
            window.addEventListener('load', pmCheckUpdate);

            function refreshData() {{
                showLoading();
                window.location.reload();
            }}
            
            function sendPaymentLink(payerId, orgId, payerName, amount, ccEmails, channel) {{
                channel = channel || 'email';
                showLoading();

                var formData = new FormData();
                formData.append('action', 'send_payment_link');
                formData.append('pid', payerId);
                formData.append('PaymentOrg', orgId);
                formData.append('PayFee', amount);
                formData.append('payerName', payerName);
                formData.append('cc_emails', ccEmails || '');
                formData.append('ProgramID', '{1}');
                formData.append('PaymentDescription', '{2}');
                formData.append('PaymentType', 'CreditCharge');
                formData.append('channel', channel);

                fetch(getPyScriptAddress(), {{
                    method: 'POST',
                    body: formData
                }})
                .then(function(response) {{
                    return response.text().then(function(text) {{
                        try {{
                            return JSON.parse(text);
                        }} catch (e) {{
                            throw new Error('Invalid JSON response: ' + text);
                        }}
                    }});
                }})
                .then(function(data) {{
                    hideLoading();
                    if (data.success) {{
                        showAlert(data.message || 'Payment link sent successfully!', 'success');
                    }} else {{
                        showAlert('Error: ' + data.message, 'danger');
                    }}
                }})
                .catch(function(error) {{
                    hideLoading();
                    showAlert('Network error: ' + error.message, 'danger');
                }});
            }}
            
            // --- Adjust balance modal ---
            // Zero out a refund-driven outstanding balance with one
            // click. Server posts an ADJ| credit equal to the
            // outstanding so the person disappears from the
            // Outstanding Balances list. Full audit trail in the
            // inline Payment History (the new line shows as
            // "Adjustment" with the reason text).
            function pmZeroOutRefund(payerId, orgId, payerName, outstanding, coregCount, coregNames) {{
                var amt = Math.abs(parseFloat(outstanding) || 0);
                if (amt <= 0) return;
                coregCount = parseInt(coregCount || 0);
                coregNames = coregNames || '';
                var confirmMsg;
                if (coregCount > 0) {{
                    confirmMsg = "Zero out this MULTI-PERSON registration?\\n\\n"
                        + "Clicking OK will clear the outstanding balance for ALL of these "
                        + "people (they share a registration that was refunded together):\\n\\n"
                        + "  - " + payerName + "\\n"
                        + "  - " + coregNames.replace(/, /g, "\\n  - ") + "\\n\\n"
                        + "An ADJ|Refund-cleared adjustment is posted per person so each "
                        + "row leaves the Outstanding Balances list. Original transactions "
                        + "remain in payment history for audit.";
                }} else {{
                    confirmMsg = "Zero out " + payerName + "'s outstanding balance of $"
                        + amt.toFixed(2) + "?\\n\\nThis posts an ADJ|Refund-cleared "
                        + "adjustment so they no longer appear here. The original "
                        + "transactions stay in the payment history for audit.";
                }}
                if (!confirm(confirmMsg)) return;
                var fd = new FormData();
                fd.append('action', 'zero_out_refund');
                fd.append('pid', payerId);
                fd.append('PaymentOrg', orgId);
                fd.append('amount', amt);
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{ return r.text(); }})
                    .then(function(txt){{
                        var d; try {{ d = JSON.parse(txt); }} catch(e) {{ showAlert('Bad response', 'danger'); return; }}
                        if (!d.success) {{ showAlert(d.message || 'Failed', 'danger'); return; }}
                        showAlert(d.message || 'Cleared', 'success');
                        setTimeout(function() {{ refreshData(); }}, 700);
                    }})
                    .catch(function(err){{ showAlert('Network error: ' + err.message, 'danger'); }});
            }}

            function openAdjustBalance(payerId, orgId, payerName, currentBalance) {{
                var modal = document.getElementById('adjustModal');
                if (!modal) return;
                document.getElementById('adj-pid').value = payerId;
                document.getElementById('adj-org').value = orgId;
                document.getElementById('adj-payer-info').textContent = payerName;
                document.getElementById('adj-current-balance').textContent =
                    '$' + (parseFloat(currentBalance) || 0).toFixed(2);
                document.getElementById('adj-direction-charge').checked = true;
                document.getElementById('adj-amount').value = '';
                document.getElementById('adj-note').value = '';
                pmAdjustPreview();
                modal.style.display = 'flex';
                setTimeout(function() {{
                    var el = document.getElementById('adj-amount');
                    if (el) el.focus();
                }}, 50);
            }}
            function closeAdjustModal() {{
                var modal = document.getElementById('adjustModal');
                if (modal) modal.style.display = 'none';
            }}
            function pmAdjustPreview() {{
                var amt = parseFloat(document.getElementById('adj-amount').value) || 0;
                var current = parseFloat((document.getElementById('adj-current-balance').textContent || '$0').replace(/[^0-9.-]/g, '')) || 0;
                var dirEl = document.querySelector('input[name="adj-direction"]:checked');
                var dir = dirEl ? dirEl.value : 'charge';
                var newBal = (dir === 'charge') ? (current + amt) : (current - amt);
                var box = document.getElementById('adj-preview');
                if (!box) return;
                if (!amt) {{ box.innerHTML = '<span style="color:#999;">Enter an amount to see the new balance.</span>'; return; }}
                var verb = (dir === 'charge') ? 'Add charge of' : 'Apply credit of';
                box.innerHTML = '<div>' + verb + ' <strong>$' + amt.toFixed(2) + '</strong></div>'
                              + '<div style="color:#666;font-size:12px;">Current: $' + current.toFixed(2)
                              + ' &rarr; <strong style="color:' + (newBal >= 0 ? '#7a5c00' : '#155724') + ';">$' + newBal.toFixed(2) + '</strong></div>';
            }}
            function submitAdjust() {{
                var payerId = document.getElementById('adj-pid').value;
                var orgId   = document.getElementById('adj-org').value;
                var amount  = document.getElementById('adj-amount').value;
                var note    = document.getElementById('adj-note').value;
                var dirEl   = document.querySelector('input[name="adj-direction"]:checked');
                var direction = dirEl ? dirEl.value : 'charge';
                var emailCb = document.getElementById('adj-email-cb');
                var sendEmail = !emailCb || emailCb.checked;
                if (!amount || parseFloat(amount) <= 0) {{ showAlert('Enter an amount greater than zero', 'warning'); return; }}
                if (!note.trim()) {{ showAlert('Please describe the reason for this adjustment', 'warning'); return; }}
                var btn = document.getElementById('adj-submit-btn');
                if (btn) {{ btn.disabled = true; btn.textContent = 'Applying...'; }}
                var fd = new FormData();
                fd.append('action', 'adjust_balance');
                fd.append('pid', payerId);
                fd.append('PaymentOrg', orgId);
                fd.append('direction', direction);
                fd.append('amount', amount);
                fd.append('note', note);
                fd.append('send_email', sendEmail ? 'true' : 'false');
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r) {{ return r.text(); }})
                    .then(function(txt) {{
                        if (btn) {{ btn.disabled = false; btn.textContent = 'Apply Adjustment'; }}
                        var d; try {{ d = JSON.parse(txt); }} catch(e) {{ showAlert('Bad response from server', 'danger'); return; }}
                        if (!d.success) {{ showAlert(d.message || 'Adjustment failed', 'danger'); return; }}
                        showAlert(d.message || 'Adjustment applied', 'success');
                        closeAdjustModal();
                        // Refresh so balances + breakdowns reflect the change.
                        setTimeout(function() {{ refreshData(); }}, 700);
                    }})
                    .catch(function(err) {{
                        if (btn) {{ btn.disabled = false; btn.textContent = 'Apply Adjustment'; }}
                        showAlert('Network error: ' + err.message, 'danger');
                    }});
            }}
            // Live-preview hooks for direction radios + amount input.
            document.addEventListener('input', function(e) {{
                if (e.target && e.target.id === 'adj-amount') pmAdjustPreview();
            }});
            document.addEventListener('change', function(e) {{
                if (e.target && e.target.name === 'adj-direction') pmAdjustPreview();
            }});

            function recordPayment(payerId, orgId, payerName, amount) {{
                var modal = document.getElementById('paymentModal');
                if (!modal) return;
                
                document.getElementById('modal-pid').value = payerId;
                document.getElementById('modal-org').value = orgId;
                document.getElementById('modal-name').value = payerName;
                document.getElementById('modal-emails').value = '';
                document.getElementById('modal-payer-info').textContent = payerName;
                document.getElementById('modal-amount-due').textContent = '$' + amount.toFixed(2);
                document.getElementById('PaidAmount').value = amount.toFixed(2);
                modal.style.display = 'flex';
            }}
            
            function closePaymentModal() {{
                var modal = document.getElementById('paymentModal');
                if (modal) modal.style.display = 'none';
            }}
            
            function submitPayment() {{
                var payerId = document.getElementById('modal-pid').value;
                var orgId = document.getElementById('modal-org').value;
                var payerName = document.getElementById('modal-name').value;
                var paymentTypeEl = document.querySelector('input[name="PaymentType"]:checked');
                var paymentType = paymentTypeEl ? paymentTypeEl.value : 'CHK|';
                var paymentDesc = document.getElementById('PaymentDescription').value;
                var paidAmount = document.getElementById('PaidAmount').value;
                
                if (!paymentDesc || !paidAmount) {{
                    showAlert('Please fill in all required fields', 'warning');
                    return;
                }}
                
                showLoading();
                closePaymentModal();
                
                var formData = new FormData();
                formData.append('action', 'record_payment');
                formData.append('pid', payerId);
                formData.append('PaymentOrg', orgId);
                formData.append('payerName', payerName);
                formData.append('PaymentType', paymentType);
                formData.append('PaymentDescription', paymentDesc);
                formData.append('PaidAmount', paidAmount);
                formData.append('ProgramID', '{1}');
                formData.append('cc_emails', '');
                
                fetch(getPyScriptAddress(), {{
                    method: 'POST',
                    body: formData
                }})
                .then(function(response) {{
                    return response.text().then(function(text) {{
                        try {{
                            return JSON.parse(text);
                        }} catch (e) {{
                            throw new Error('Invalid JSON response: ' + text);
                        }}
                    }});
                }})
                .then(function(data) {{
                    hideLoading();
                    if (data.success) {{
                        showAlert('Payment recorded successfully!', 'success');
                        setTimeout(function() {{ refreshData(); }}, 1500);
                    }} else {{
                        showAlert('Error: ' + data.message, 'danger');
                    }}
                }})
                .catch(function(error) {{
                    hideLoading();
                    showAlert('Network error: ' + error.message, 'danger');
                }});
            }}
            
            function viewTransactions(payerId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=transactions&PeopleId=' + payerId;
            }}

            // Inline expand on the Outstanding cell. Shows the chain of
            // charges + payments that produced the current balance,
            // without leaving the page.
            function pmToggleChargeDetails(anchor, peopleId, orgId) {{
                var row = anchor.closest('tr');
                var detailRow = row && row.nextElementSibling;
                var caret = anchor.querySelector('.pm-charge-toggle');
                if (!detailRow || !detailRow.classList.contains('pm-charge-row')) return;
                // Check current computed display state -- not just the
                // inline style, which can be '' once we've toggled and
                // the row inherits its default (table-row).
                var isOpen = window.getComputedStyle(detailRow).display !== 'none';
                if (isOpen) {{
                    detailRow.style.display = 'none';
                    if (caret) caret.style.transform = 'rotate(0deg)';
                    return;
                }}
                detailRow.style.display = 'table-row';
                if (caret) caret.style.transform = 'rotate(90deg)';
                var body = detailRow.querySelector('.pm-charge-body');
                if (!body) return;
                // If we've already loaded for this row, leave the cache
                // alone (toggle just shows/hides).
                if (detailRow.getAttribute('data-loaded') === '1') return;
                body.innerHTML = '<div style="padding:6px 0;color:#999;"><i class="fa fa-spinner fa-spin"></i> Loading charges...</div>';
                var fd = new FormData();
                fd.append('action', 'charge_details');
                fd.append('people_id', peopleId);
                fd.append('org_id', orgId);
                fetch(getPyScriptAddress(), {{ method: 'POST', body: fd }})
                    .then(function(r){{ return r.text(); }})
                    .then(function(txt){{
                        var d; try {{ d = JSON.parse(txt); }} catch(e) {{
                            body.innerHTML = '<div style="color:#a94442;">Bad response from server.</div>';
                            return;
                        }}
                        if (!d.success) {{
                            body.innerHTML = '<div style="color:#a94442;">' + pmEsc(d.message || 'Failed') + '</div>';
                            return;
                        }}
                        var rows = d.charges || [];
                        if (!rows.length) {{
                            body.innerHTML = '<div style="color:#999;">No transactions on file for this involvement.</div>';
                        }} else {{
                            // Server returns chronological so running
                            // balances are correct. Reverse for display
                            // so the newest activity sits at the top --
                            // typical accounting-ledger style.
                            rows.reverse();
                            var html = '<table style="width:100%;font-size:12px;border-collapse:collapse;">';
                            html += '<thead><tr style="text-align:left;color:#666;">';
                            html += '<th style="padding:4px 8px;font-weight:600;">Date</th>';
                            html += '<th style="padding:4px 8px;font-weight:600;">Kind</th>';
                            html += '<th style="padding:4px 8px;font-weight:600;">Note / Description</th>';
                            html += '<th style="padding:4px 8px;font-weight:600;text-align:right;">Charge</th>';
                            html += '<th style="padding:4px 8px;font-weight:600;text-align:right;">Payment</th>';
                            html += '<th style="padding:4px 8px;font-weight:600;text-align:right;">Balance after</th>';
                            html += '</tr></thead><tbody>';
                            for (var i = 0; i < rows.length; i++) {{
                                var r = rows[i];
                                var amt = Math.abs(parseFloat(r.amount) || 0);
                                var bal = parseFloat(r.runningBalance) || 0;
                                var balStr = (bal < 0 ? '-' : '') + pmFmtMoney(bal);
                                var balColor = bal > 0 ? '#7a5c00' : (bal < 0 ? '#3c763d' : '#666');
                                // Refunds get a red pill so they stand out --
                                // staff often miss them and chase a balance
                                // that was actually money returned to the
                                // payer. Other charges = amber, payments =
                                // green.
                                var kindStr = String(r.kind || '');
                                var isRefund = kindStr.indexOf('Refund') >= 0;
                                var kindBg, kindFg;
                                if (isRefund) {{ kindBg = '#f8d7da'; kindFg = '#a8071a'; }}
                                else if (r.isPayment) {{ kindBg = '#d4edda'; kindFg = '#155724'; }}
                                else {{ kindBg = '#fff4ce'; kindFg = '#7a5c00'; }}
                                html += '<tr style="border-top:1px solid #e1e5eb;">';
                                html += '<td style="padding:4px 8px;white-space:nowrap;">' + pmEsc(r.tranDate || '') + '</td>';
                                html += '<td style="padding:4px 8px;"><span style="background:' + kindBg + ';color:' + kindFg + ';padding:1px 6px;border-radius:8px;font-size:11px;font-weight:600;">' + pmEsc(r.kind || '') + '</span></td>';
                                // Description cell -- show the staff note
                                // first (italicized), then the involvement
                                // context underneath in a smaller dim font.
                                // Tooltip carries the raw Message in case
                                // someone needs to see the full string.
                                var noteText = (r.note || '').trim();
                                var descText = (r.description || '').trim();
                                var noteCell = '';
                                if (noteText) noteCell += '<div style="color:#444;font-style:italic;" title="' + pmEsc(r.message || '') + '">' + pmEsc(noteText) + '</div>';
                                if (descText) noteCell += '<div style="font-size:11px;color:#999;margin-top:1px;">' + pmEsc(descText) + '</div>';
                                if (!noteCell) noteCell = '<span style="color:#bbb;">&mdash;</span>';
                                html += '<td style="padding:4px 8px;">' + noteCell + '</td>';
                                html += '<td style="padding:4px 8px;text-align:right;color:#7a5c00;font-weight:600;">' + (r.isPayment ? '' : pmFmtMoney(amt)) + '</td>';
                                html += '<td style="padding:4px 8px;text-align:right;color:#155724;font-weight:600;">' + (r.isPayment ? pmFmtMoney(amt) : '') + '</td>';
                                html += '<td style="padding:4px 8px;text-align:right;font-weight:700;color:' + balColor + ';">' + balStr + '</td>';
                                html += '</tr>';
                            }}
                            html += '</tbody></table>';
                            var ending = parseFloat(d.endingBalance) || 0;
                            html += '<div style="margin-top:6px;font-size:11px;color:#666;font-style:italic;">' + rows.length + ' transaction(s), newest first -- current balance ' + (ending < 0 ? '-' : '') + pmFmtMoney(ending) + '. Positive = the person owes the church; "Balance after" is the running total in chronological order.</div>';
                            body.innerHTML = html;
                        }}
                        detailRow.setAttribute('data-loaded', '1');
                    }})
                    .catch(function(err){{
                        body.innerHTML = '<div style="color:#a94442;">Network error: ' + pmEsc(err.message) + '</div>';
                    }});
            }}
            
            function viewEmails(payerId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=emails&PeopleId=' + payerId;
            }}
            
            function previewEmail(messageId, peopleId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=email_preview&messageId=' + messageId + '&PeopleId=' + peopleId;
            }}
            
            function resendEmail(messageId, peopleId) {{
                if (!confirm('Are you sure you want to resend this email?')) {{
                    return;
                }}
                
                showLoading();
                
                var formData = new FormData();
                formData.append('action', 'resend_email');
                formData.append('messageId', messageId);
                formData.append('PeopleId', peopleId);
                
                fetch(getPyScriptAddress(), {{
                    method: 'POST',
                    body: formData
                }})
                .then(function(response) {{
                    return response.text().then(function(text) {{
                        try {{
                            return JSON.parse(text);
                        }} catch (e) {{
                            throw new Error('Invalid JSON response: ' + text);
                        }}
                    }});
                }})
                .then(function(data) {{
                    hideLoading();
                    if (data.success) {{
                        showAlert('Email resent successfully!', 'success');
                    }} else {{
                        showAlert('Error: ' + data.message, 'danger');
                    }}
                }})
                .catch(function(error) {{
                    hideLoading();
                    showAlert('Network error: ' + error.message, 'danger');
                }});
            }}
            </script>
            
            <style>
            .pm-container {{
                max-width: 1200px;
                margin: 20px auto;
                padding: 0 20px;
                font-family: 'Segoe UI', Arial, sans-serif;
            }}
            .pm-header {{
                background: #f8f9fa;
                padding: 15px 20px;
                border-radius: 8px;
                margin-bottom: 20px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border: 1px solid #dee2e6;
            }}
            .pm-header h3 {{
                margin: 0;
                color: #495057;
                font-size: 1.5rem;
            }}
            .pm-actions {{
                display: flex;
                gap: 10px;
            }}
            .pm-search {{
                margin-bottom: 20px;
            }}
            .pm-search .form-control {{
                max-width: 400px;
                padding: 8px 12px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-size: 14px;
            }}
            .pm-content {{
                background: white;
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                border: 1px solid #dee2e6;
            }}
            .pm-table {{
                width: 100%;
                border-collapse: collapse;
                margin: 0;
            }}
            .pm-table th {{
                background: #343a40;
                color: white;
                padding: 12px;
                text-align: left;
                font-weight: 600;
                border: none;
            }}
            .pm-table td {{
                padding: 12px;
                border-bottom: 1px solid #dee2e6;
                vertical-align: top;
            }}
            .pm-table tbody tr:hover {{
                background-color: #f8f9fa;
            }}
            .pm-table tfoot {{
                background: #f8f9fa;
                font-weight: 600;
            }}
            .pm-table tfoot th {{
                background: #e9ecef;
                color: #495057;
                text-align: left;
            }}
            .pm-table tfoot th.pm-currency {{
                text-align: right;
            }}
            .pm-table tfoot th.pm-center {{
                text-align: center;
            }}
            .pm-currency {{
                text-align: right;
                font-weight: 600;
                color: #28a745;
            }}
            .pm-positive {{
                color: #28a745;
            }}
            .pm-negative {{
                color: #dc3545;
            }}
            .pm-center {{
                text-align: center;
            }}
            .pm-contact {{
                color: #6c757d;
                font-size: 0.9em;
                margin-top: 4px;
            }}
            .pm-muted {{
                color: #6c757d;
            }}
            .pm-btn-group {{
                display: flex;
                flex-direction: column;
                gap: 4px;
                min-width: 120px;
            }}
            .btn {{
                padding: 6px 12px;
                border: 1px solid;
                border-radius: 4px;
                cursor: pointer;
                text-decoration: none;
                display: inline-block;
                text-align: center;
                font-size: 14px;
                line-height: 1.4;
                background: white;
                transition: all 0.2s;
            }}
            .btn:hover {{
                opacity: 0.8;
                transform: translateY(-1px);
            }}
            .btn-primary {{
                background: #007bff;
                border-color: #007bff;
                color: white;
            }}
            .btn-secondary {{
                background: #6c757d;
                border-color: #6c757d;
                color: white;
            }}
            .btn-success {{
                background: #28a745;
                border-color: #28a745;
                color: white;
            }}
            .btn-outline-primary {{
                border-color: #007bff;
                color: #007bff;
            }}
            .btn-outline-success {{
                border-color: #28a745;
                color: #28a745;
            }}
            .btn-outline-info {{
                border-color: #17a2b8;
                color: #17a2b8;
            }}
            .btn-outline-secondary {{
                border-color: #6c757d;
                color: #6c757d;
            }}
            .btn-sm {{
                padding: 4px 8px;
                font-size: 12px;
            }}
            .fa {{
                margin-right: 4px;
            }}
            .pm-loading {{
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(255,255,255,0.8);
                display: none;
                align-items: center;
                justify-content: center;
                z-index: 9999;
            }}
            .pm-loading-spinner {{
                text-align: center;
                padding: 20px;
                background: white;
                border-radius: 8px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            }}
            .pm-hidden {{
                display: none !important;
            }}
            .email-meta {{
                margin-bottom: 30px;
                padding: 20px;
                background: #f8f9fa;
                border-radius: 8px;
            }}
            .email-meta h4 {{
                margin-top: 0;
                color: #495057;
            }}
            .email-body {{
                padding: 20px;
            }}
            .email-body h4 {{
                margin-top: 0;
                color: #495057;
            }}
            .email-content {{
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 20px;
                background: white;
                min-height: 300px;
            }}
            .alert {{
                padding: 15px;
                margin-bottom: 20px;
                border: 1px solid transparent;
                border-radius: 4px;
            }}
            .alert-warning {{
                color: #856404;
                background-color: #fff3cd;
                border-color: #ffeaa7;
            }}
            .alert-danger {{
                color: #721c24;
                background-color: #f8d7da;
                border-color: #f5c6cb;
            }}
            .alert-success {{
                color: #155724;
                background-color: #d4edda;
                border-color: #c3e6cb;
            }}
            @keyframes spin {{
                0% {{ transform: rotate(0deg); }}
                100% {{ transform: rotate(360deg); }}
            }}
            @media (max-width: 768px) {{
                .pm-container {{
                    padding: 0 10px;
                }}
                .pm-header {{
                    flex-direction: column;
                    gap: 10px;
                    text-align: center;
                }}
                .pm-table {{
                    font-size: 14px;
                }}
                .pm-table th, .pm-table td {{
                    padding: 8px;
                }}
                .pm-btn-group {{
                    flex-direction: row;
                    flex-wrap: wrap;
                }}
            }}
            </style>
            
            <div class="pm-loading" id="pmLoading">
                <div class="pm-loading-spinner">
                    <div style="width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid #007bff; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto;"></div>
                    <div style="margin-top: 10px;">Processing...</div>
                </div>
            </div>
            
            <!-- Alert container for messages -->
            <div id="alertContainer" style="position: fixed; top: 20px; right: 20px; z-index: 10001;"></div>

            <!-- Auto-update banner (populated by pmCheckUpdate on load) -->
            <div id="pmUpdateBanner" style="max-width:1200px;margin:8px auto 0;padding:0 14px;"></div>

            <!-- Setup banner (server-rendered; empty when fully configured) -->
            {6}

            <!-- Toolbar -->
            <div style="max-width:1200px;margin:8px auto 12px;padding:8px 14px;
                        display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                <div style="flex:1;font-size:13px;color:#666;">
                    Payment Manager <span style="color:#999;font-size:11px;">v{3}</span>
                </div>
                <button class="btn btn-primary" onclick="viewPrograms()" style="font-size:12px;font-weight:700;" title="Home -- outstanding balances by program">
                    <i class="fa fa-home"></i> Outstanding Balances
                </button>
                <button class="btn btn-secondary" onclick="viewReceipts()" style="font-size:12px;">
                    <i class="fa fa-receipt"></i> Receipts
                </button>
                <button class="btn btn-secondary" onclick="openHelp()" title="What does this tool do?" style="font-size:12px;">
                    ? Help
                </button>
                <button class="btn btn-secondary" onclick="openSettings()" title="Email senders, templates, defaults" style="font-size:12px;">
                    &#9881; Settings
                </button>
            </div>

            {0}

            <!-- Customize-receipt modal -->
            <div id="pmCustomModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:10003;align-items:center;justify-content:center;padding:20px;">
                <div style="background:#fff;border-radius:10px;max-width:1100px;width:100%;max-height:90vh;overflow:hidden;display:flex;flex-direction:column;">
                    <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #eee;padding:14px 22px;">
                        <div>
                            <h3 style="margin:0;color:#1f4e79;font-size:18px;"><i class="fa fa-edit"></i> Customize Receipt</h3>
                            <div style="font-size:12px;color:#666;margin-top:2px;">Payer: <strong id="pmCustomTo">&mdash;</strong></div>
                        </div>
                        <button onclick="pmCustomClose()" style="background:none;border:0;font-size:24px;color:#999;cursor:pointer;">&times;</button>
                    </div>
                    <div style="padding:14px 22px 4px;border-bottom:1px solid #eee;">
                        <div style="display:flex;gap:12px;margin-bottom:8px;align-items:flex-end;">
                            <div style="flex:1;">
                                <label style="display:block;font-size:12px;color:#666;font-weight:600;margin-bottom:3px;">Send to <span id="pmCustomToHint" style="font-weight:400;color:#999;">(defaults to email on file)</span></label>
                                <input type="email" id="pmCustomToEmail" placeholder="recipient@example.com" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
                            </div>
                            <button type="button" onclick="pmCustomResetEmail()" class="btn btn-sm" style="font-size:11px;padding:5px 10px;height:fit-content;">Use email on file</button>
                        </div>
                        <label style="display:block;font-size:12px;color:#666;font-weight:600;margin-bottom:3px;">Subject</label>
                        <input type="text" id="pmCustomSubject" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
                    </div>
                    <div style="display:flex;gap:14px;padding:14px 22px;flex:1;overflow:hidden;">
                        <div style="flex:1;display:flex;flex-direction:column;min-width:0;">
                            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                                <label style="font-size:12px;color:#666;font-weight:600;">Personal Note (optional)</label>
                                <span style="font-size:11px;color:#999;">Plain text &mdash; appears above the receipt</span>
                            </div>
                            <textarea id="pmCustomNote" oninput="pmCustomRefreshPreview()" placeholder="Add a note here -- e.g., 'Thank you for your generous gift to summer camp!' Leave blank to send the standard receipt." style="flex:1;width:100%;min-height:380px;font-family:Segoe UI,Tahoma,sans-serif;font-size:13px;padding:10px;border:1px solid #ccc;border-radius:4px;resize:vertical;line-height:1.4;"></textarea>
                            <div style="font-size:11px;color:#999;margin-top:4px;font-style:italic;">Tip: blank lines become paragraph breaks. The receipt below stays the same; your note is added above it.</div>
                        </div>
                        <div style="flex:1;display:flex;flex-direction:column;min-width:0;">
                            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                                <label style="font-size:12px;color:#666;font-weight:600;">Preview</label>
                                <span style="font-size:11px;color:#999;">Live as you type</span>
                            </div>
                            <div id="pmCustomPreview" style="flex:1;border:1px solid #ccc;border-radius:4px;padding:14px;overflow:auto;background:#fff;font-family:Segoe UI,Tahoma,sans-serif;font-size:13px;"></div>
                        </div>
                    </div>
                    <div style="display:flex;justify-content:flex-end;gap:8px;padding:12px 22px;border-top:1px solid #eee;background:#f8f9fa;">
                        <button onclick="pmCustomClose()" class="btn btn-secondary">Cancel</button>
                        <button onclick="pmCustomPrint()" class="btn"><i class="fa fa-print"></i> Print Edited</button>
                        <button id="pmCustomSendBtn" onclick="pmCustomSend()" class="btn btn-primary"><i class="fa fa-paper-plane"></i> Send Customized Receipt</button>
                    </div>
                </div>
            </div>

            <!-- Settings modal -->
            <div id="pmSettingsModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:10002;align-items:center;justify-content:center;padding:20px;">
                <div style="background:#fff;border-radius:10px;max-width:820px;width:100%;max-height:88vh;overflow-y:auto;padding:18px 22px;">
                    <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #eee;padding-bottom:10px;margin-bottom:14px;">
                        <h3 style="margin:0;color:#1f4e79;">&#9881; Payment Manager Settings</h3>
                        <button onclick="closeSettings()" style="background:none;border:0;font-size:24px;color:#999;cursor:pointer;">&times;</button>
                    </div>
                    <div style="font-size:13px;color:#333;">
                        <div style="font-weight:700;color:#1f4e79;margin-bottom:6px;">Default Email Sender</div>
                        <p style="color:#666;font-size:12px;margin:0 0 8px;">Used for every program unless an override below matches. Replaces the old hardcoded values.</p>
                        <div style="display:grid;grid-template-columns:1fr 2fr;gap:6px 12px;margin-bottom:14px;">
                            <label>Sender PeopleId</label><input type="text" id="pmSetSenderId" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                            <label>From Email</label><input type="text" id="pmSetSenderEmail" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                            <label>From Name (Alias)</label><input type="text" id="pmSetSenderAlias" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                            <label>Email Title</label><input type="text" id="pmSetSenderTitle" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                            <label>Reply Phone</label><input type="text" id="pmSetSenderPhone" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                        </div>

                        <div style="font-weight:700;color:#1f4e79;margin-bottom:6px;">Email Templates</div>
                        <p style="color:#666;font-size:12px;margin:0 0 8px;">HTML Content names. Falls back to built-in defaults if the named template doesn't exist.</p>
                        <div style="display:grid;grid-template-columns:1fr 2fr;gap:6px 12px;margin-bottom:14px;">
                            <label>Payment Confirmation</label><input type="text" id="pmSetTplConfirm" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                            <label>Payment Notification</label><input type="text" id="pmSetTplNotify" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                            <label>Subject Prefix</label><input type="text" id="pmSetSubjectPrefix" style="padding:5px 8px;border:1px solid #ccc;border-radius:4px;">
                        </div>

                        <div style="font-weight:700;color:#1f4e79;margin:18px 0 6px;">Email Template Bodies</div>
                        <p style="color:#666;font-size:12px;margin:0 0 8px;">The editors below are pre-filled with a sensible default as a starter. Edit the HTML, then click <strong>Save</strong> to write it to TouchPoint's HTML Content under the names configured above. <strong>Reset</strong> reloads the built-in default into the editor. <strong>Preview</strong> opens a popup with sample merge values.</p>
                        <details style="margin:0 0 10px;background:#f6f8fa;border:1px solid #e0e0e0;border-radius:6px;padding:8px 12px;font-size:12px;">
                            <summary style="cursor:pointer;font-weight:600;color:#1f4e79;">Merge fields you can use</summary>
                            <table style="width:100%;margin-top:8px;border-collapse:collapse;">
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{name}}</code></td><td>Payer's name</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{chargeNotes}}</code></td><td>Payment details (type, amount, date) &mdash; includes the involvement + registered names</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{orgName}}</code></td><td>Involvement name (separately, in case you want it on its own line)</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{registrants}}</code></td><td>Comma-separated list of who was registered</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{note}}</code></td><td>Staff note attached to the payment (the text after the CHK| / CSH| / FEE| / ADJ| prefix)</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{previousDue}}</code></td><td>Previous balance</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{totalDue}}</code></td><td>Total now due (notification template)</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{newTotalDue}}</code></td><td>New balance after payment (confirmation template)</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{paylink}}</code></td><td>Payment link URL (notification template)</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{sender_alias}}</code></td><td>Sender display name</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{sender_phone}}</code></td><td>Sender phone</td></tr>
                                <tr><td style="padding:2px 8px 2px 0;white-space:nowrap;"><code>{{sender_email}}</code></td><td>Sender email</td></tr>
                            </table>
                        </details>

                        <div style="margin-bottom:12px;">
                            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                                <label style="font-weight:600;color:#1f4e79;">Payment Confirmation <span id="pmTplConfirmName" style="font-weight:400;color:#666;font-family:monospace;font-size:11px;"></span> <span id="pmTplConfirmBadge" style="font-weight:400;color:#999;font-size:11px;font-style:italic;"></span></label>
                                <div>
                                    <button class="btn btn-sm" onclick="pmPreviewTemplate('confirmation')" style="font-size:11px;padding:3px 10px;">Preview</button>
                                    <button class="btn btn-sm" onclick="pmResetTemplate('confirmation')" style="font-size:11px;padding:3px 10px;">Reset</button>
                                    <button class="btn btn-sm btn-primary" onclick="pmSaveTemplate('confirmation')" style="font-size:11px;padding:3px 10px;">Save</button>
                                </div>
                            </div>
                            <textarea id="pmTplConfirmBody" rows="8" style="width:100%;font-family:Menlo,Consolas,monospace;font-size:12px;padding:8px;border:1px solid #ccc;border-radius:4px;"></textarea>
                            <p style="font-size:11px;color:#666;margin:4px 0 0;font-style:italic;">Sent after a check / cash payment is recorded. Also used by the Receipts tab for duplicate reprints.</p>
                        </div>

                        <div style="margin-bottom:18px;">
                            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                                <label style="font-weight:600;color:#1f4e79;">Payment Notification <span id="pmTplNotifyName" style="font-weight:400;color:#666;font-family:monospace;font-size:11px;"></span> <span id="pmTplNotifyBadge" style="font-weight:400;color:#999;font-size:11px;font-style:italic;"></span></label>
                                <div>
                                    <button class="btn btn-sm" onclick="pmPreviewTemplate('notification')" style="font-size:11px;padding:3px 10px;">Preview</button>
                                    <button class="btn btn-sm" onclick="pmResetTemplate('notification')" style="font-size:11px;padding:3px 10px;">Reset</button>
                                    <button class="btn btn-sm btn-primary" onclick="pmSaveTemplate('notification')" style="font-size:11px;padding:3px 10px;">Save</button>
                                </div>
                            </div>
                            <textarea id="pmTplNotifyBody" rows="8" style="width:100%;font-family:Menlo,Consolas,monospace;font-size:12px;padding:8px;border:1px solid #ccc;border-radius:4px;"></textarea>
                            <p style="font-size:11px;color:#666;margin:4px 0 0;font-style:italic;">Sent when a payment link is generated for someone. Recipient gets a link to pay online.</p>
                        </div>

                        <div style="font-weight:700;color:#1f4e79;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center;">
                            <span>Per-Program Sender Overrides</span>
                            <button onclick="pmAddSender()" class="btn btn-sm btn-primary" style="font-size:11px;padding:3px 10px;">+ Add Override</button>
                        </div>
                        <p style="color:#666;font-size:12px;margin:0 0 8px;">Optional. When set, payments from this program use these sender values instead of the Default above.</p>
                        <table class="pm-table" style="font-size:12px;">
                            <thead><tr><th>Program ID</th><th>Sender ID</th><th>From Email</th><th>Alias</th><th>Title</th><th></th></tr></thead>
                            <tbody id="pmSendersTbody"></tbody>
                        </table>
                    </div>
                    <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:18px;padding-top:12px;border-top:1px solid #eee;">
                        <button onclick="closeSettings()" class="btn btn-secondary">Cancel</button>
                        <button onclick="pmSaveSettings()" class="btn btn-primary">Save Settings</button>
                    </div>
                </div>
            </div>

            <!-- Help modal -->
            <div id="pmHelpModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:10002;align-items:center;justify-content:center;padding:20px;">
                <div style="background:#fff;border-radius:10px;max-width:720px;width:100%;max-height:88vh;overflow-y:auto;padding:18px 22px;">
                    <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #eee;padding-bottom:10px;margin-bottom:14px;">
                        <h3 style="margin:0;color:#1f4e79;">Payment Manager Help</h3>
                        <button onclick="closeHelp()" style="background:none;border:0;font-size:24px;color:#999;cursor:pointer;">&times;</button>
                    </div>
                    <div style="font-size:13px;color:#1e293b;line-height:1.55;">
                        <h4 style="color:#1f4e79;margin:14px 0 6px;">Programs &rarr; Divisions &rarr; Payers</h4>
                        <p>Drill from the top-level <strong>Outstanding Balances</strong> view into a program's divisions, then into individual payers with unpaid balances. From a payer row you can <em>Send Payment Link</em>, <em>Record Payment</em>, view <em>Emails</em>, or open the full <em>Transaction</em> history.</p>
                        <h4 style="color:#1f4e79;margin:14px 0 6px;">&#128201; Receipts (new)</h4>
                        <p>Find past check or cash payments by date range. For each, <strong>Print</strong> opens a popup-window receipt (good for the church files) and <strong>Email</strong> re-sends the original-template receipt to the payer with a "Duplicate Receipt" banner so it's clearly not a new charge.</p>
                        <h4 style="color:#1f4e79;margin:14px 0 6px;">&#9881; Settings (new)</h4>
                        <p>All the old hardcoded constants (default sender, template names, subject prefix) now live here. Per-program sender overrides replace the legacy <code>PM_EmailSenders</code> dict (which is auto-imported on first save).</p>
                        <h4 style="color:#1f4e79;margin:14px 0 6px;">Auto-update</h4>
                        <p>When a newer version is published a banner appears at the top of the page. One click pulls the latest code; your settings are stored separately and survive the update.</p>
                    </div>
                    <div style="display:flex;justify-content:flex-end;margin-top:18px;padding-top:12px;border-top:1px solid #eee;">
                        <button onclick="closeHelp()" class="btn btn-primary">Got it</button>
                    </div>
                </div>
            </div>
            
            <!-- Payment Modal -->
            <div id="paymentModal" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 10000; align-items: center; justify-content: center;">
                <div style="background: white; padding: 20px; border-radius: 8px; max-width: 500px; width: 90%;">
                    <h4>Record Payment</h4>
                    <div id="paymentForm">
                        <input type="hidden" id="modal-pid">
                        <input type="hidden" id="modal-org">
                        <input type="hidden" id="modal-name">
                        <input type="hidden" id="modal-emails">
                        
                        <div style="margin-bottom: 15px;">
                            <label>Payer:</label>
                            <div id="modal-payer-info" style="font-weight: bold;"></div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Amount Due:</label>
                            <div id="modal-amount-due" style="color: green; font-weight: bold;"></div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Payment Type:</label><br>
                            <label><input type="radio" name="PaymentType" value="CHK|" checked> Check</label>
                            <label style="margin-left: 15px;"><input type="radio" name="PaymentType" value="CSH|"> Cash</label>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Description:</label>
                            <input type="text" id="PaymentDescription" style="width: 100%; padding: 8px; border: 1px solid #ccc;" required>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Payment Amount:</label>
                            <input type="number" id="PaidAmount" step="0.01" min="0" style="width: 100%; padding: 8px; border: 1px solid #ccc;" required>
                        </div>
                        
                        <div style="text-align: right;">
                            <button type="button" onclick="closePaymentModal()" class="btn btn-secondary">Cancel</button>
                            <button type="button" onclick="submitPayment()" class="btn btn-primary" style="margin-left: 10px;">Record Payment</button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Adjust balance modal -->
            <div id="adjustModal" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:10001;align-items:center;justify-content:center;">
                <div style="background:#fff;padding:24px;border-radius:8px;max-width:480px;width:90%;">
                    <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #eee;padding-bottom:10px;margin-bottom:16px;">
                        <h3 style="margin:0;color:#1f4e79;font-size:18px;"><i class="fa fa-pencil"></i> Adjust Balance</h3>
                        <button type="button" onclick="closeAdjustModal()" style="background:none;border:0;font-size:24px;color:#999;cursor:pointer;">&times;</button>
                    </div>
                    <input type="hidden" id="adj-pid">
                    <input type="hidden" id="adj-org">
                    <div style="margin-bottom:12px;">
                        <label style="color:#666;font-size:12px;">Payer</label>
                        <div id="adj-payer-info" style="font-weight:700;font-size:14px;"></div>
                    </div>
                    <div style="margin-bottom:14px;">
                        <label style="color:#666;font-size:12px;">Current Balance</label>
                        <div id="adj-current-balance" style="color:#7a5c00;font-weight:700;font-size:16px;"></div>
                    </div>
                    <div style="margin-bottom:14px;">
                        <label style="color:#666;font-size:12px;font-weight:600;display:block;margin-bottom:6px;">Adjustment type</label>
                        <label style="display:flex;align-items:center;gap:8px;padding:6px 10px;border:1px solid #ddd;border-radius:4px;margin-bottom:6px;cursor:pointer;">
                            <input type="radio" name="adj-direction" id="adj-direction-charge" value="charge" checked>
                            <span><strong>Add charge</strong> <span style="color:#666;font-size:12px;">&mdash; increases what they owe</span></span>
                        </label>
                        <label style="display:flex;align-items:center;gap:8px;padding:6px 10px;border:1px solid #ddd;border-radius:4px;cursor:pointer;">
                            <input type="radio" name="adj-direction" id="adj-direction-credit" value="credit">
                            <span><strong>Apply credit</strong> <span style="color:#666;font-size:12px;">&mdash; reduces what they owe</span></span>
                        </label>
                    </div>
                    <div style="margin-bottom:14px;">
                        <label for="adj-amount" style="color:#666;font-size:12px;font-weight:600;">Amount (positive)</label>
                        <input type="number" id="adj-amount" min="0.01" step="0.01" placeholder="0.00" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;font-size:14px;">
                    </div>
                    <div style="margin-bottom:14px;">
                        <label for="adj-note" style="color:#666;font-size:12px;font-weight:600;">Reason / description</label>
                        <input type="text" id="adj-note" placeholder="e.g., Scholarship approved by Pastor; partial refund; etc." style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;font-size:13px;">
                        <div style="font-size:11px;color:#999;margin-top:3px;font-style:italic;">Saved as ADJ|&lt;reason&gt; on the transaction.</div>
                    </div>
                    <div id="adj-preview" style="background:#f8f9fa;border:1px solid #e1e5eb;border-radius:6px;padding:10px 14px;margin-bottom:14px;font-size:13px;"></div>
                    <label style="display:flex;align-items:center;gap:8px;font-size:13px;color:#444;margin-bottom:14px;cursor:pointer;padding:8px 10px;background:#eef5fb;border:1px solid #cfe3ff;border-radius:6px;">
                        <input type="checkbox" id="adj-email-cb" checked>
                        <span><strong>Email the payer</strong> about this adjustment <span style="color:#666;font-size:11px;">&mdash; includes the reason, new balance, and a pay link if a balance remains</span></span>
                    </label>
                    <div style="text-align:right;">
                        <button type="button" onclick="closeAdjustModal()" class="btn btn-secondary">Cancel</button>
                        <button type="button" id="adj-submit-btn" onclick="submitAdjust()" class="btn btn-primary" style="margin-left:10px;">Apply Adjustment</button>
                    </div>
                </div>
            </div>

            <script>
            // Setup page functionality after DOM is loaded
            document.addEventListener('DOMContentLoaded', function() {{
                // Setup search functionality
                var searchInput = document.getElementById('searchInput');
                if (searchInput) {{
                    searchInput.addEventListener('input', function() {{
                        var filter = this.value.toLowerCase();
                        var table = document.getElementById('payersTable');
                        if (!table) return;
                        
                        var rows = table.querySelectorAll('tbody tr');
                        for (var i = 0; i < rows.length; i++) {{
                            var text = rows[i].textContent.toLowerCase();
                            rows[i].style.display = text.indexOf(filter) > -1 ? '' : 'none';
                        }}
                    }});
                }}
                
                // Setup transaction search
                var transactionSearch = document.getElementById('transactionSearch');
                if (transactionSearch) {{
                    transactionSearch.addEventListener('input', function() {{
                        var filter = this.value.toLowerCase();
                        var table = document.getElementById('transactionsTable');
                        if (!table) return;
                        
                        var rows = table.querySelectorAll('tbody tr');
                        for (var i = 0; i < rows.length; i++) {{
                            var text = rows[i].textContent.toLowerCase();
                            rows[i].style.display = text.indexOf(filter) > -1 ? '' : 'none';
                        }}
                    }});
                }}
                
                // Setup modal functionality
                var paymentModal = document.getElementById('paymentModal');
                if (paymentModal) {{
                    paymentModal.addEventListener('click', function(e) {{
                        if (e.target === this) {{
                            closePaymentModal();
                        }}
                    }});
                }}
                
                // Setup form submission
                var paymentForm = document.getElementById('paymentForm');
                if (paymentForm) {{
                    paymentForm.addEventListener('keypress', function(e) {{
                        if (e.key === 'Enter') {{
                            e.preventDefault();
                            submitPayment();
                        }}
                    }});
                }}
                
                hideLoading();
            }});
            </script>
            """.format(content, self.program_id, PAYMENT_LINK_DESCRIPTION, APP_VERSION, DC_SCRIPT_ID, DC_API_BASE, self._build_setup_banner_html())
            
        def render(self):
            """Main render method - determines which view to show and handles POST requests"""
            try:
                # Handle POST requests (AJAX actions)
                if hasattr(model.Data, 'action'):
                    action = str(model.Data.action)
                    
                    # Check if this is a POST request for AJAX actions
                    POST_ACTIONS = [
                        'send_payment_link', 'record_payment', 'resend_email',
                        'find_receipts', 'reprint_receipt', 'send_custom_receipt',
                        'charge_details', 'adjust_balance', 'zero_out_refund',
                        'load_settings', 'save_settings',
                        'save_template', 'reset_template',
                        'apply_update',
                    ]
                    if action in POST_ACTIONS:
                        if action == 'send_payment_link':     return self.process_payment_link()
                        if action == 'record_payment':        return self.process_payment_record()
                        if action == 'resend_email':          return self.process_resend_email()
                        if action == 'find_receipts':         return self.process_find_receipts()
                        if action == 'reprint_receipt':       return self.process_reprint_receipt()
                        if action == 'send_custom_receipt':   return self.process_send_custom_receipt()
                        if action == 'charge_details':        return self.process_charge_details()
                        if action == 'adjust_balance':        return self.process_adjust_balance()
                        if action == 'zero_out_refund':       return self.process_zero_out_refund()
                        if action == 'load_settings':         return self.process_load_settings()
                        if action == 'save_settings':         return self.process_save_settings()
                        if action == 'save_template':         return self.process_save_template()
                        if action == 'reset_template':        return self.process_reset_template()
                        if action == 'apply_update':          return self.process_apply_update()
                
                # Handle GET requests (page views)
                content = ""
                
                if self.current_action == 'divisions':
                    program_id = str(getattr(model.Data, 'ProgramID', self.program_id))
                    content = self.render_divisions_view(program_id)
                elif self.current_action == 'payers':
                    org_id = getattr(model.Data, 'OrganizationId', None)
                    program_id = getattr(model.Data, 'ProgramID', None)
                    if org_id:
                        org_id = str(org_id)
                    if program_id:
                        program_id = str(program_id)
                    content = self.render_payers_view(org_id, program_id)
                elif self.current_action == 'emails':
                    people_id = str(getattr(model.Data, 'PeopleId', ''))
                    if people_id:
                        content = self.render_email_history(people_id)
                    else:
                        content = "<div class='alert alert-warning'>People ID is required for email history</div>"
                elif self.current_action == 'email_preview':
                    message_id = str(getattr(model.Data, 'messageId', ''))
                    people_id = str(getattr(model.Data, 'PeopleId', ''))
                    if message_id and people_id:
                        content = self.render_email_preview(message_id, people_id)
                    else:
                        content = "<div class='alert alert-warning'>Message ID and People ID are required for email preview</div>"
                elif self.current_action == 'transactions':
                    people_id = str(getattr(model.Data, 'PeopleId', ''))
                    if people_id:
                        content = self.render_transaction_history(people_id)
                    else:
                        content = "<div class='alert alert-warning'>People ID is required for transaction history</div>"
                elif self.current_action == 'receipts':
                    content = self.render_receipts_view()
                else:
                    content = self.render_programs_view()
                
                return self.render_page_structure(content)
                
            except Exception as e:
                return self.render_error_page(str(e))
            
        def render_error_page(self, error_message):
            """Render an error page with details"""
            import traceback
            error_details = traceback.format_exc()
            
            return """
            <div style="max-width: 800px; margin: 50px auto; padding: 20px;">
                <div style="background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 8px; padding: 20px; color: #721c24;">
                    <h3><i class="fa fa-exclamation-triangle"></i> Error</h3>
                    <p>An error occurred while processing your request:</p>
                    <p><strong>{0}</strong></p>
                    <hr>
                    <details>
                        <summary>Technical Details</summary>
                        <pre style="background: white; padding: 10px; border-radius: 4px; overflow: auto;">{1}</pre>
                    </details>
                    <div style="margin-top: 20px;">
                        <button onclick="history.go(-1)" class="btn btn-primary">Go Back</button>
                        <button onclick="location.reload()" class="btn btn-secondary" style="margin-left: 10px;">Retry</button>
                    </div>
                </div>
            </div>
            """.format(error_message, error_details)

    # Main execution
    payment_manager = ModernPaymentManager()
    
    # Set page title
    model.Title = PAGE_TITLE
    
    # Print the rendered page
    print(payment_manager.render())

except Exception as e:
    # Print any errors
    import traceback
    print("<div style='max-width: 800px; margin: 50px auto; padding: 20px;'>")
    print("<div style='background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 8px; padding: 20px; color: #721c24;'>")
    print("<h2><i class='fa fa-exclamation-triangle'></i> System Error</h2>")
    print("<p>An unexpected error occurred: <strong>{0}</strong></p>".format(str(e)))
    print("<hr>")
    print("<details><summary>Technical Details</summary>")
    print("<pre style='background: white; padding: 10px; border-radius: 4px;'>")
    traceback.print_exc()
    print("</pre></details>")
    print("<div style='margin-top: 20px;'>")
    print("<button onclick='history.go(-1)' style='padding: 8px 16px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer;'>Go Back</button>")
    print("<button onclick='location.reload()' style='padding: 8px 16px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px;'>Retry</button>")
    print("</div></div></div>")
