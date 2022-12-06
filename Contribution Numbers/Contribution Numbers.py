#Roles=rpt_BusinessAdmin

#sql statement is based on fiscal year.  Ours runs Oct - Sept, so fiscal year is 3 as it's 3 months from the new year. 
#If you want to it to be calendar year, just change it to 0 and update yeartype to nothing or something else you want it to pre-fix
model.Header = 'Contribution Numbers'

fiscalmonth = '3'
yeartype = 'FY'

sqlFundReport = '''
Select
  DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) AS [Year],
  Sum(c.ContributionAmount) AS [Contributed],
  Round(Sum(c.ContributionAmount)/Count(Distinct c.ContributionID),2) AS [AverageGift],
  Count(c.ContributionID) AS [Gifts], 
  Count(Distinct c.PeopleID) AS [UniqueGivers],
  Count(c.ContributionID)/Count(Distinct c.PeopleID) AS [AvgNumofGifts],
  Count(CASE WHEN c.ContributionAmount Between 10000 AND 99999 THEN c.ContributionID END) AS [Gifts10kto99k],
  Count(CASE WHEN c.ContributionAmount >= 100000 THEN c.ContributionID END) AS [Gifts100kPlus]
FROM Contribution c
LEFT JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE Year(c.ContributionDate) >= (DATEPART(Year,getdate())-10)
 Group By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) ---Year(c.ContributionDate) 
 Order By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate)))---Year(c.ContributionDate)
'''

template = '''
<table ; width="800px" border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Year</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Contributed</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Avg Gift</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5># of Gifts</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Unique Givers</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Avg # of Gifts</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>10k-99k</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>100k+</h5></b></td>
    </tr>
    {{#each fundreport}}
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><a href="https://myfbch.com/PyScript/ContributionFundYear?p1={{Year}}" target="_blank">''' + yeartype + '''{{Year}}</a></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt Contributed 'N2'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt AverageGift 'N2'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt Gifts 'N0'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt UniqueGivers 'N0'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{AvgNumofGifts}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Gifts10kto99k}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Gifts100kPlus}}</td>
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
