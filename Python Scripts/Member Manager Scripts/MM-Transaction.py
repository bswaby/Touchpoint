#title=FMC Transaction

ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName
model.Header = ProgramName + ' Transaction'

from datetime import datetime, date, time

def ProcessGet():
    form = model.DynamicData()

    html = model.Content(ProgramName + 'Transaction')
    model.Form = model.RenderTemplate(html)
    
def GetContTypeId(code):
    contTypeDict = {'c': 1, 'n': 9, 'p': 8, 's': 20, 'g': 10}
    return contTypeDict.get(code, 1) # defaut to ContributionTypeId = 1 if code not found
    
def ProcessPost():
    startdate = model.ParseDate(model.Data.startdate)
    enddate = model.ParseDate(model.Data.enddate)
    sqlstartdate = datetime.strptime(str(startdate), "%m/%d/%Y %I:%M:%S %p")
    sqlenddate = datetime.strptime(str(enddate), "%m/%d/%Y %I:%M:%S %p")

    sqlBalance = '''
    SELECT 
        dbo.TransactionSummary.PeopleId, 
        dbo.TransactionList.Name, 
        dbo.TransactionList.TransactionDate, 
        dbo.TransactionList.Description, 
        dbo.People.Name AS [CompletedBy], 
        dbo.TransactionList.coupon, 
        dbo.TransactionList.IsAdjustment, 
        dbo.TransactionList.moneytran, 
        dbo.TransactionList.Message, 
        dbo.TransactionList.TotalPayment, 
        dbo.TransactionList.Payment,
        dbo.TransactionList.Id 
    FROM 
        dbo.TransactionList 
    INNER JOIN 
        dbo.TransactionSummary 
    ON 
        ( 
            dbo.TransactionList.OriginalId = dbo.TransactionSummary.RegId) 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.TransactionList.LoginPeopleId = dbo.People.PeopleId) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.TransactionList.OrgId = Organizations_alias1.OrganizationId) 
    INNER JOIN 
        dbo.Division 
    ON 
        ( 
            Organizations_alias1.DivisionId = dbo.Division.Id) 
    WHERE 
        dbo.Division.ProgId = ''' + ProgramID + ''' 
    AND dbo.TransactionList.TotalPayment <> 0
    AND dbo.TransactionList.TransactionDate Between ''' + "'" + str(sqlstartdate) + "'" + ''' AND ''' + "'" + str(sqlenddate) + "'" + '''
    ORDER BY 
        dbo.TransactionList.TransactionDate DESC;
    '''
    print sqlBalance
    sqlTotalCoupons = '''
    SELECT 
        Sum(dbo.TransactionList.TotalPayment) AS [TotalPayment]
    FROM 
        dbo.TransactionList 
    INNER JOIN 
        dbo.TransactionSummary 
    ON 
        ( 
            dbo.TransactionList.OriginalId = dbo.TransactionSummary.RegId) 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.TransactionList.LoginPeopleId = dbo.People.PeopleId) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.TransactionList.OrgId = Organizations_alias1.OrganizationId) 
    INNER JOIN 
        dbo.Division 
    ON 
        ( 
            Organizations_alias1.DivisionId = dbo.Division.Id) 
    WHERE 
        dbo.Division.ProgId = ''' + ProgramID + ''' 
    AND dbo.TransactionList.TotalPayment <> 0 
    AND dbo.TransactionList.coupon = 1
    AND dbo.TransactionList.TransactionDate Between ''' + "'" + str(sqlstartdate) + "'" + ''' AND ''' + "'" + str(sqlenddate) + "'" + ''';
    '''
    
    sqlTotalChecks = '''
    SELECT 
        Sum(dbo.TransactionList.TotalPayment) AS [TotalPayment]
    FROM 
        dbo.TransactionList 
    INNER JOIN 
        dbo.TransactionSummary 
    ON 
        ( 
            dbo.TransactionList.OriginalId = dbo.TransactionSummary.RegId) 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.TransactionList.LoginPeopleId = dbo.People.PeopleId) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.TransactionList.OrgId = Organizations_alias1.OrganizationId) 
    INNER JOIN 
        dbo.Division 
    ON 
        ( 
            Organizations_alias1.DivisionId = dbo.Division.Id) 
    WHERE 
        dbo.Division.ProgId = ''' + ProgramID + ''' 
    AND dbo.TransactionList.TotalPayment <> 0 
    AND dbo.TransactionList.Message Like 'CHK%'
    AND dbo.TransactionList.TransactionDate Between ''' + "'" + str(sqlstartdate) + "'" + ''' AND ''' + "'" + str(sqlenddate) + "'" + ''';
    '''
    
    sqlTotalCash = '''
    SELECT 
        Sum(dbo.TransactionList.TotalPayment) AS [TotalPayment]
    FROM 
        dbo.TransactionList 
    INNER JOIN 
        dbo.TransactionSummary 
    ON 
        ( 
            dbo.TransactionList.OriginalId = dbo.TransactionSummary.RegId) 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.TransactionList.LoginPeopleId = dbo.People.PeopleId) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.TransactionList.OrgId = Organizations_alias1.OrganizationId) 
    INNER JOIN 
        dbo.Division 
    ON 
        ( 
            Organizations_alias1.DivisionId = dbo.Division.Id) 
    WHERE 
        dbo.Division.ProgId = ''' + ProgramID + ''' 
    AND dbo.TransactionList.TotalPayment <> 0 
    AND dbo.TransactionList.Message Like 'CSH%'
    AND dbo.TransactionList.TransactionDate Between ''' + "'" + str(sqlstartdate) + "'" + ''' AND ''' + "'" + str(sqlenddate) + "'" + ''';
    '''
    sqlTotalAdjustments = '''
    SELECT 
        Sum(dbo.TransactionList.TotalPayment) AS [TotalPayment]
    FROM 
        dbo.TransactionList 
    INNER JOIN 
        dbo.TransactionSummary 
    ON 
        ( 
            dbo.TransactionList.OriginalId = dbo.TransactionSummary.RegId) 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.TransactionList.LoginPeopleId = dbo.People.PeopleId) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.TransactionList.OrgId = Organizations_alias1.OrganizationId) 
    INNER JOIN 
        dbo.Division 
    ON 
        ( 
            Organizations_alias1.DivisionId = dbo.Division.Id) 
    WHERE 
        dbo.Division.ProgId = ''' + ProgramID + ''' 
    AND dbo.TransactionList.TotalPayment <> 0 
    AND dbo.TransactionList.Message Like 'ADJ%'
    AND dbo.TransactionList.TransactionDate Between ''' + "'" + str(sqlstartdate) + "'" + ''' AND ''' + "'" + str(sqlenddate) + "'" + ''';
    '''  
    sqlTotalOther = '''
    SELECT 
        Sum(dbo.TransactionList.TotalPayment) AS [TotalPayment]
    FROM 
        dbo.TransactionList 
    INNER JOIN 
        dbo.TransactionSummary 
    ON 
        ( 
            dbo.TransactionList.OriginalId = dbo.TransactionSummary.RegId) 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.TransactionList.LoginPeopleId = dbo.People.PeopleId) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.TransactionList.OrgId = Organizations_alias1.OrganizationId) 
    INNER JOIN 
        dbo.Division 
    ON 
        ( 
            Organizations_alias1.DivisionId = dbo.Division.Id) 
    WHERE 
        dbo.Division.ProgId = ''' + ProgramID + ''' 
    AND dbo.TransactionList.TotalPayment <> 0 
    AND dbo.TransactionList.Message Not Like 'CHK%'
    AND dbo.TransactionList.Message Not Like 'CSH%'
    AND dbo.TransactionList.TransactionDate Between ''' + "'" + str(sqlstartdate) + "'" + ''' AND ''' + "'" + str(sqlenddate) + "'" + ''';
    '''
    template = '''
        <h3>Totals</h3>
        
          <div class="stat-panel-inner-container">
                <div class="stat-panel">
                    <span class="stat-title">COUPONS</span>
                    <span class="stat-figure">{{#each totalcoupons}}{{FmtMoney TotalPayment}}{{/each}}</span>
                </div>
                <div class="stat-panel">
                    <span class="stat-title">CHECKS</span>
                    <span class="stat-figure">{{#each totalchecks}}{{FmtMoney TotalPayment}}{{/each}}</span>
                </div>
                <div class="stat-panel">
                    <span class="stat-title">CASH</span>
                    <span class="stat-figure">{{#each totalcash}}{{FmtMoney TotalPayment}}{{/each}}</span>
                </div>
                <div class="stat-panel">
                    <span class="stat-title">ADJUSTMENTS</span>
                    <span class="stat-figure">{{#each totaladjustments}}{{FmtMoney TotalPayment}}{{/each}}</span>
                </div>
                <div class="stat-panel">
                    <span class="stat-title">OTHER</span>
                    <span class="stat-figure">{{#each totalother}}{{FmtMoney TotalPayment}}{{/each}}</span>
                </div>
                <div style="clear: both;">
                </div>
          </div>

        <h3>Transactions</h3>
        <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
        <script type="text/javascript">
    
          google.charts.load('current', {'packages':['table']});
          google.charts.setOnLoadCallback(drawTable);
    
          function drawTable() {
            var data = new google.visualization.DataTable();
            data.addColumn('string', 'Name');
            data.addColumn('string', 'Transaction');
            data.addColumn('string', 'Description');
            data.addColumn('string', 'Completed By');
            data.addColumn('string', 'Coupon');
            data.addColumn('string', 'Adjustment');
            data.addColumn('string', 'MoneyTran');
            data.addColumn('string', 'Message');
            data.addColumn('number', 'Total Due');
            
            data.addRows([
                {{#each balancereport}}
                    [
                    '<a href="https://fbchvillebak.tpsdb.com/Person2/{{PeopleId}}#tab-registrations" target="_blank">{{Name}}</a>', 
                    '<a href="https://fbchvillebak.tpsdb.com/Transactions/{{Id}}" target="_blank">{{TransactionDate}}</a>',
                    '{{Description}}',
                    '{{CompletedBy}}',
                    '{{coupon}}',
                    '{{IsAdjustment}}',
                    '{{moneytran}}',
                    '{{Message}}',
                    {v: {{TotalPayment}},  f: '{{FmtMoney TotalPayment}}'},
                    ],
                {{/each}}
            ]);
            
            var table = new google.visualization.Table(document.getElementById('table_div'));
            table.draw(data, {showRowNumber: false, alternatingRowStyle: true, allowHtml: true, width: '100%'});
          }
        </script>
        <div id='table_div' style='width: 1000px;'></div>
    '''

    Data.totalcoupons = q.QuerySql(sqlTotalCoupons)
    Data.totalchecks = q.QuerySql(sqlTotalChecks)
    Data.totalcash = q.QuerySql(sqlTotalCash)
    Data.totaladjustments = q.QuerySql(sqlTotalAdjustments)
    Data.totalother = q.QuerySql(sqlTotalOther)
    Data.balancereport = q.QuerySql(sqlBalance)
    NMReport = model.RenderTemplate(template)
    print(NMReport)
    #print sqlTotalOther

if model.HttpMethod == 'get':
    ProcessGet()
elif model.HttpMethod == 'post':
    ProcessPost()