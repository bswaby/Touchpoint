#Roles=rpt_BusinessAdmin
model.Header = '@P1 Fund Breakdown'


sqlFundReport = '''
SELECT 
    DISTINCT dbo.ContributionFund.FundName, 
    SUM(dbo.Contribution.ContributionAmount) AS Total
--    SUM(CASE WHEN DATEPART(Year,DATEADD(month,3,(dbo.Contribution.ContributionDate))) = @P1 THEN dbo.Contribution.ContributionAmount END) AS [a2020]
--    SUM(CASE WHEN DATEPART(Year,DATEADD(month,3,(dbo.Contribution.ContributionDate))) = {{y2}} THEN dbo.Contribution.ContributionAmount END) AS [b2021],
--    SUM(CASE WHEN DATEPART(Year,DATEADD(month,3,(dbo.Contribution.ContributionDate))) = {{y3}} THEN dbo.Contribution.ContributionAmount END) AS [c2022] 
FROM 
    dbo.Contribution 
INNER JOIN 
    dbo.ContributionFund 
ON 
    ( 
        dbo.Contribution.FundId = dbo.ContributionFund.FundId) 
WHERE 
    DATEPART(Year,DATEADD(month,3,(dbo.Contribution.ContributionDate))) = @P1
GROUP BY 
    dbo.ContributionFund.FundName;
'''   

template = '''
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
      google.charts.load('current', {'packages':['table']});
      google.charts.setOnLoadCallback(drawTable);

      function drawTable() {
        var data = new google.visualization.DataTable();
        data.addColumn('string', 'Fund');
        data.addColumn('number', 'Total');

        data.addRows([
            {{#each fundreport}}
                [
                '{{FundName}}',
                {v: {{Total}},  f: '{{FmtMoney Total}}'}
                ],
            {{/each}}
        ]);
        
        var table = new google.visualization.Table(document.getElementById('table_wklybudget'));
        table.draw(data, {showRowNumber: false, alternatingRowStyle: true, width: '100%', height: '100%'});
      }
    </script>
    <div id='table_wklybudget' style='width: 400px; height: 800px;'></div>
'''

Data.fundreport = q.QuerySql(sqlFundReport)
NMReport = model.RenderTemplate(template)
print(NMReport)
