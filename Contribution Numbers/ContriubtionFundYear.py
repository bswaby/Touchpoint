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
<table ; width="400px" border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    {{#each fundreport}}
    <tr style="{{Bold}}">
            <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FundName}}</td>
            <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt Total 'N2'}}</td>
            <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt c2020 'N2'}}</td>
    </tr>
    {{/each}}
    </tbody>
</table>
<br />
<br />
'''


Data.fundreport = q.QuerySql(sqlFundReport)
NMReport = model.RenderTemplate(template)
print(NMReport)
