#Roles=Edit
# Written By: Ben Swaby (TPxi Software, LLC)
# Email: bswaby@fbchtn.org                                                                                                      
# Website: https://tpxisoftware.com
# GitHub: https://github.com/bswaby/Touchpoint  (50+ free tools)                                                                
# ----------------------------------------------------------------                                                              
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to                                                                
# support continued development, check out:                                                                                     
#
# DisplayCache(TM) - church digital signage that integrates with TouchPoint(R)                                                  
# https://displaycache.com                                
#
# TPxi Go(TM) - your church contacts, wherever you work.
# Look up anyone in TouchPoint(R), log calls and emails from Outlook                                                            
# or your phone. No tab switching, no lost context.
# https://tpxigo.com                                                                                                            
# ----------------------------------------------------------------
# Description: Search transactions by the last 4 digits of a
#              credit card or ACH account number. Returns matching
#              transactions with person, organization, payment
#              method, and amount details. Defaults to the last
#              90 days but supports custom date ranges.
# ---------------------------------------------------------------


model.Header = 'Transaction Search'

import json

# Handle AJAX search
if model.HttpMethod == 'post' and hasattr(Data, 'action') and str(Data.action) == 'search':
    last4 = str(Data.last4).strip() if hasattr(Data, 'last4') else ''
    date_from = str(Data.date_from).strip() if hasattr(Data, 'date_from') else ''
    date_to = str(Data.date_to).strip() if hasattr(Data, 'date_to') else ''

    # Validate last 4 digits
    if not last4 or len(last4) != 4 or not last4.isdigit():
        print json.dumps({'success': False, 'message': 'Please enter exactly 4 digits.'})
    else:
        # Build date filter
        date_filter = ''
        if date_from:
            date_filter += " AND t.TransactionDate >= '{0}'".format(date_from.replace("'", ""))
        if date_to:
            date_filter += " AND t.TransactionDate <= '{0} 23:59:59'".format(date_to.replace("'", ""))

        # Default to last 90 days if no date range specified
        if not date_from and not date_to:
            date_filter = " AND t.TransactionDate >= DATEADD(day, -90, GETDATE())"

        sql = '''
            SELECT TOP 200
                t.Id,
                t.TransactionDate,
                t.TransactionId,
                t.Message,
                t.LastFourCC,
                t.LastFourACH,
                t.PaymentType,
                CASE
                    WHEN t.LastFourCC = '{0}' THEN 'Credit Card (****' + t.LastFourCC + ')'
                    WHEN t.LastFourACH = '{0}' THEN 'ACH (****' + t.LastFourACH + ')'
                    ELSE ISNULL(t.PaymentType, 'Unknown')
                END as PayMethod,
                FORMAT(t.amt, 'C') AS Amount,
                t.amt as RawAmt,
                t.TransactionGateway,
                t.Name as TranName,
                t.First,
                t.Last,
                t.Emails as TranEmail,
                t.Phone as TranPhone,
                tp.PeopleId,
                tp.OrgId,
                p.Name2,
                p.EmailAddress,
                p.CellPhone,
                org.OrganizationName,
                t.Description
            FROM [Transaction] t
            LEFT JOIN [TransactionPeople] tp ON t.OriginalId = tp.Id
            LEFT JOIN Organizations org ON t.OrgId = org.OrganizationId
            LEFT JOIN People p ON p.PeopleId = tp.PeopleId
            WHERE (t.LastFourCC = '{0}' OR t.LastFourACH = '{0}')
                AND t.amt <> 0
                {1}
            ORDER BY t.TransactionDate DESC
        '''.format(last4, date_filter)

        results = q.QuerySql(sql)
        rows = []
        for r in results:
            # Use person record if linked, otherwise fall back to transaction fields
            name = str(r.Name2) if r.Name2 else str(r.TranName) if r.TranName else 'Unknown'
            email = str(r.EmailAddress) if r.EmailAddress else str(r.TranEmail) if r.TranEmail else ''
            phone = str(r.CellPhone) if r.CellPhone else str(r.TranPhone) if r.TranPhone else ''

            rows.append({
                'Id': r.Id,
                'TransactionDate': str(r.TransactionDate) if r.TransactionDate else '',
                'TransactionId': str(r.TransactionId) if r.TransactionId else '',
                'PayMethod': str(r.PayMethod) if r.PayMethod else '',
                'Description': str(r.Description) if r.Description else '',
                'Amount': str(r.Amount) if r.Amount else '$0.00',
                'RawAmt': float(r.RawAmt) if r.RawAmt else 0,
                'Gateway': str(r.TransactionGateway) if r.TransactionGateway else '',
                'LastFourCC': str(r.LastFourCC) if r.LastFourCC else '',
                'LastFourACH': str(r.LastFourACH) if r.LastFourACH else '',
                'PaymentType': str(r.PaymentType) if r.PaymentType else '',
                'PeopleId': r.PeopleId if r.PeopleId else 0,
                'Name': name,
                'Email': email,
                'Phone': phone,
                'OrgId': r.OrgId if r.OrgId else 0,
                'OrgName': str(r.OrganizationName) if r.OrganizationName else ''
            })

        print json.dumps({'success': True, 'count': len(rows), 'results': rows})

else:
    # Render search form
    model.Script = '''
    $(function() {
        // Default date range: last 90 days
        var today = new Date();
        var past = new Date();
        past.setDate(past.getDate() - 90);
        $('#date_to').val(today.toISOString().split('T')[0]);
        $('#date_from').val(past.toISOString().split('T')[0]);

        // Allow Enter key to trigger search
        $('#last4').on('keypress', function(e) {
            if (e.which === 13) { doSearch(); }
        });

        $('#last4').focus();
    });

    function doSearch() {
        var last4 = $('#last4').val().trim();
        if (last4.length !== 4 || !/^\\d{4}$/.test(last4)) {
            alert('Please enter exactly 4 digits.');
            return;
        }

        $('#search-btn').prop('disabled', true).text('Searching...');
        $('#results-area').html('<div style="text-align:center;padding:40px;"><i class="fa fa-spinner fa-spin fa-2x"></i><br>Searching transactions...</div>');

        var scriptName = window.location.pathname.split('/').pop().split('?')[0];

        $.ajax({
            url: '/PyScriptForm/' + scriptName,
            type: 'POST',
            data: {
                action: 'search',
                last4: last4,
                date_from: $('#date_from').val(),
                date_to: $('#date_to').val()
            },
            success: function(response) {
                try {
                    var data = JSON.parse(response);
                    if (!data.success) {
                        $('#results-area').html('<div class="alert alert-warning">' + data.message + '</div>');
                    } else if (data.count === 0) {
                        $('#results-area').html('<div class="alert alert-info">No transactions found matching last 4 digits: <strong>' + last4 + '</strong></div>');
                    } else {
                        renderResults(data.results, data.count, last4);
                    }
                } catch(e) {
                    $('#results-area').html('<div class="alert alert-danger">Error parsing response.</div>');
                }
                $('#search-btn').prop('disabled', false).text('Search');
            },
            error: function() {
                $('#results-area').html('<div class="alert alert-danger">Search failed. Please try again.</div>');
                $('#search-btn').prop('disabled', false).text('Search');
            }
        });
    }

    function renderResults(results, count, last4) {
        var total = 0;
        for (var i = 0; i < results.length; i++) {
            total += results[i].RawAmt;
        }

        var html = '<div style="margin-bottom:15px;">';
        html += '<span class="label label-primary" style="font-size:14px;padding:6px 12px;">' + count + ' transaction(s) found</span> ';
        html += '<span class="label label-default" style="font-size:14px;padding:6px 12px;">Last 4: ' + last4 + '</span> ';
        html += '<span class="label label-success" style="font-size:14px;padding:6px 12px;">Total: $' + total.toFixed(2) + '</span>';
        html += '</div>';

        html += '<div class="table-responsive"><table class="table table-striped table-hover table-condensed">';
        html += '<thead><tr>';
        html += '<th>Date</th><th>Person</th><th>Organization</th><th>Payment Method</th><th>Description</th><th style="text-align:right">Amount</th><th>Actions</th>';
        html += '</tr></thead><tbody>';

        for (var i = 0; i < results.length; i++) {
            var r = results[i];
            var dateStr = r.TransactionDate ? r.TransactionDate.split(' ')[0] : '';
            var personLink = r.PeopleId > 0
                ? '<a href="/Person2/' + r.PeopleId + '#tab-registrations" target="_blank">' + r.Name + '</a>'
                : r.Name;
            var orgLink = r.OrgId > 0
                ? '<a href="/Organization/' + r.OrgId + '" target="_blank">' + r.OrgName + '</a>'
                : r.OrgName;
            var rowClass = r.RawAmt < 0 ? 'style="color:#c9302c;"' : '';

            html += '<tr ' + rowClass + '>';
            html += '<td>' + dateStr + '</td>';
            html += '<td>' + personLink;
            if (r.Email) html += '<br><small class="text-muted">' + r.Email + '</small>';
            if (r.Phone) html += '<br><small class="text-muted">' + r.Phone + '</small>';
            html += '</td>';
            html += '<td>' + orgLink + '</td>';
            html += '<td>' + r.PayMethod + '</td>';
            html += '<td><small>' + r.Description + '</small></td>';
            html += '<td style="text-align:right;font-weight:bold;">' + r.Amount + '</td>';
            html += '<td><a href="/Transactions/' + r.Id + '" target="_blank" class="btn btn-xs btn-default" title="View Transaction"><i class="fa fa-external-link"></i></a></td>';
            html += '</tr>';
        }

        html += '</tbody></table></div>';

        if (count >= 200) {
            html += '<div class="alert alert-warning">Results limited to 200 rows. Narrow your date range for more specific results.</div>';
        }

        $('#results-area').html(html);
    }
    '''

    model.Form = '''
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <style>
        .search-container {
            max-width: 900px;
            margin: 0 auto;
        }
        .search-box {
            background: #f5f5f5;
            border: 1px solid #ddd;
            border-radius: 6px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .search-box h4 {
            margin-top: 0;
            color: #333;
        }
        .input-group-lg input {
            font-size: 18px;
            letter-spacing: 8px;
            text-align: center;
            font-weight: bold;
        }
        .help-text {
            font-size: 12px;
            color: #888;
            margin-top: 5px;
        }
    </style>

    <div class="search-container">
        <div class="search-box">
            <h4><i class="fa fa-search"></i> Search Transactions by Last 4 Digits</h4>
            <p class="text-muted">Enter the last 4 digits of an ACH account or credit card number to find matching transactions.</p>

            <div class="row">
                <div class="col-sm-4">
                    <label>Last 4 Digits (ACH/CC)</label>
                    <input type="text" id="last4" class="form-control input-lg" maxlength="4" placeholder="0000" pattern="[0-9]{4}" autocomplete="off">
                    <p class="help-text">Searches LastFourCC and LastFourACH fields</p>
                </div>
                <div class="col-sm-3">
                    <label>From Date</label>
                    <input type="date" id="date_from" class="form-control">
                </div>
                <div class="col-sm-3">
                    <label>To Date</label>
                    <input type="date" id="date_to" class="form-control">
                </div>
                <div class="col-sm-2">
                    <label>&nbsp;</label>
                    <button id="search-btn" class="btn btn-primary btn-block" onclick="doSearch()">
                        <i class="fa fa-search"></i> Search
                    </button>
                </div>
            </div>
        </div>

        <div id="results-area"></div>
    </div>
    '''
