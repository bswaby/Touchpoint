model.Header = 'Involvement Balance'

sqlDueTotals = '''
SELECT 
  SUM(CASE WHEN (dbo.TransactionSummary.TranDate <= GetDate() 
             AND dbo.TransactionSummary.TranDate >= DateAdd(mm,-1, GetDate())) 
        THEN dbo.TransactionSummary.TotDue END) AS [CurrentDue],
  SUM(CASE WHEN (dbo.TransactionSummary.TranDate <= DateAdd(mm,-1, GetDate()) 
              AND dbo.TransactionSummary.TranDate >= DateAdd(mm,-3, GetDate())) 
         THEN dbo.TransactionSummary.TotDue END) AS [PastDue30-90],
  SUM(CASE WHEN (dbo.TransactionSummary.TranDate <= DateAdd(mm,-3, GetDate()) 
              AND dbo.TransactionSummary.TranDate >= DateAdd(mm,-6, GetDate())) 
         THEN dbo.TransactionSummary.TotDue END) AS [PastDue90-180] 
FROM 
    dbo.Organizations Organizations_alias1  
INNER JOIN 
    dbo.TransactionSummary 
ON 
    ( 
        Organizations_alias1.OrganizationId = dbo.TransactionSummary.OrganizationId) 

WHERE 
    Organizations_alias1.OrganizationId = @P1
'''

sqlBalance = '''
SELECT 
    dbo.People.FamilyId, 
    dbo.People.PeopleId, 
    dbo.People.Name2 AS [MemberName], 
    dbo.People.CellPhone,
    dbo.People.HomePhone,
    dbo.TransactionSummary.TranDate, 
    Organizations_alias1.OrganizationName,
    dbo.MemberTags.Name AS [SubGroup],
    dbo.TransactionSummary.TotDue,
    dbo.TransactionSummary.TotPaid
FROM dbo.Organizations Organizations_alias1 
INNER JOIN 
    dbo.TransactionSummary 
ON 
    ( 
        Organizations_alias1.OrganizationId = dbo.TransactionSummary.OrganizationId) 
LEFT OUTER JOIN 
    dbo.People 
ON 
    ( 
        dbo.TransactionSummary.PeopleId = dbo.People.PeopleId) 
LEFT JOIN
    dbo.OrgMemMemTags
ON
    (dbo.OrgMemMemTags.PeopleId = dbo.People.PeopleId) AND (dbo.OrgMemMemTags.OrgId = Organizations_alias1.OrganizationId)
LEFT JOIN 
    dbo.MemberTags
ON
    (dbo.MemberTags.Id = dbo.OrgMemMemTags.MemberTagId)
WHERE 
Organizations_alias1.OrganizationId = @P1 
Order By     dbo.TransactionSummary.TotDue Desc, dbo.People.Name2;
'''

sqlOrganizations = '''Select OrganizationName, OrganizationId FROM Organizations Order by OrganizationName'''


header = '''
<script src ="https://code.jquery.com/jquery-3.5.1.min.js"></script>                 
<script src ="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"> </script>  
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.0.0/jquery.min.js"></script>
<!-- jQuery Modal -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css" />
'''

body = '''
  <table>
	   {{#each duetotals}}
		  <tr><td width="200px"><h4>CurrentDue</h4></td><td align="right"><h4>{{FmtMoney CurrentDue}}</h4></td></tr>
		  <tr><td width="200px"><h4>PastDue 30-90</h4></td><td align="right"><h4>{{FmtMoney PastDue30-90}}</h4></td></tr>
		  <tr><td width="200px"><h4>PastDue 90-180</h4></td><td align="right"><h4>{{FmtMoney PastDue90-180}}</h4></td></tr>
		{{/each}}
  </table>
    <br>
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">

      google.charts.load('current', {'packages':['table']});
      google.charts.setOnLoadCallback(drawTable);

      function drawTable() {
        var data = new google.visualization.DataTable();
        data.addColumn('number', 'FamilyId');
        data.addColumn('string', 'PeopleId');
        data.addColumn('string', 'Name');
        data.addColumn('string', 'TranDate');
        data.addColumn('string', 'Involvement (subgroup)');
        data.addColumn('number', 'Due');
        data.addColumn('number', 'Paid');
        data.addColumn('string', '');
        data.addRows([
            {{#each balancereport}}
                [
                {{FamilyId}}, 
                '{{PeopleId}}',
                '{{MemberName}}', 
                '{{TranDate}}',
                '{{OrganizationName}} ({{SubGroup}})',
                {v: {{TotDue}},  f: '{{FmtMoney TotDue}}'},
                {v: {{TotPaid}},  f: '{{FmtMoney TotPaid}}'},
                '<a href="/Person2/{{PeopleId}}#tab-current" target="_blank"><i class="fa fa-info-circle" aria-hidden="true"></i></a>',
                ],
            {{/each}}
        ]);
        
        var table = new google.visualization.Table(document.getElementById('table_div'));
        table.draw(data, {showRowNumber: false, alternatingRowStyle: true, allowHtml: true, width: '100%', height: '100%'});
      }
    </script>
    <div id='table_div' style='width: 900px;'></div>
'''
footer = '''
<script>
    $(function(){
      // bind change event to select
      $('#org_select').on('change', function () {
          var url = $(this).val(); // get selected value
          if (url) { // require a URL
              window.location = url; // redirect
          }
          return false;
      });
    });
</script>
'''

print model.RenderTemplate(header)

print '''<select id="org_select"><option value="Ben?p1=0"></option>'''

for a in q.QuerySql(sqlOrganizations):
    print '<option value="Ben?p1={1}">{0}</option>'.format(a.OrganizationName, a.OrganizationId)

print '''</select>'''

Data.duetotals = q.QuerySql(sqlDueTotals)
Data.balancereport = q.QuerySql(sqlBalance)
print model.RenderTemplate(body)
print model.RenderTemplate(footer)
