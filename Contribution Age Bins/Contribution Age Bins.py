#Roles=rpt_BusinessAdmin

#Update fiscal month to your need.  If you want calendar year set to 0.  If you want August as your start month, count backwards..
#so in the case of an August fiscal year, you would put in 5.

model.Header = 'Contribution By Age Bin'

fiscalmonth = '3'

sqlgifts = '''
SELECT  
  DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) AS [UniqueContributorsbyYear],
  Sum(CASE WHEN p.Age >= 0 AND p.Age < 10 THEN c.ContributionAmount END) AS [ca],
  Sum(CASE WHEN p.Age >= 10 AND p.Age < 20 THEN c.ContributionAmount END) AS [cb],
  Sum(CASE WHEN p.Age >= 20 AND p.Age < 30 THEN c.ContributionAmount END) AS [cc],
  Sum(CASE WHEN p.Age >= 30 AND p.Age < 40 THEN c.ContributionAmount END) AS [cd],
  Sum(CASE WHEN p.Age >= 40 AND p.Age < 50 THEN c.ContributionAmount END) AS [ce],
  Sum(CASE WHEN p.Age >= 50 AND p.Age < 60 THEN c.ContributionAmount END) AS [cf],
  Sum(CASE WHEN p.Age >= 60 AND p.Age < 70 THEN c.ContributionAmount END) AS [cg],
  Sum(CASE WHEN p.Age >= 70 AND p.Age < 80 THEN c.ContributionAmount END) AS [ch],
  Sum(CASE WHEN p.Age >= 80 AND p.Age < 90 THEN c.ContributionAmount END) AS [ci],
  Sum(CASE WHEN p.Age >= 90 THEN c.ContributionAmount END) AS [cj],
  Sum(c.ContributionAmount) AS [ck],
  Count(Distinct CASE WHEN p.Age >= 0 AND p.Age < 10 THEN p.PeopleID END) AS [pa],
  Count(Distinct CASE WHEN p.Age >= 10 AND p.Age < 20 THEN p.PeopleID END) AS [pb],
  Count(Distinct CASE WHEN p.Age >= 20 AND p.Age < 30 THEN p.PeopleID END) AS [pc],
  Count(Distinct CASE WHEN p.Age >= 30 AND p.Age < 40 THEN p.PeopleID END) AS [pd],
  Count(Distinct CASE WHEN p.Age >= 40 AND p.Age < 50 THEN p.PeopleID END) AS [pe],
  Count(Distinct CASE WHEN p.Age >= 50 AND p.Age < 60 THEN p.PeopleID END) AS [pf],
  Count(Distinct CASE WHEN p.Age >= 60 AND p.Age < 70 THEN p.PeopleID END) AS [pg],
  Count(Distinct CASE WHEN p.Age >= 70 AND p.Age < 80 THEN p.PeopleID END) AS [ph],
  Count(Distinct CASE WHEN p.Age >= 80 AND p.Age < 90 THEN p.PeopleID END) AS [pi],
  Count(Distinct CASE WHEN p.Age >= 90 THEN p.PeopleID END) AS [pj],
  Count(Distinct p.PeopleID) AS [pk]
FROM Contribution c
INNER JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) >= (DATEPART(Year,getdate())-10)
 Group By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate)))
 Order By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate)))
'''

sqlgiftstotal = '''
SELECT  
  Sum(CASE WHEN p.Age >= 0 AND p.Age < 10 THEN c.ContributionAmount END) AS [cl],
  Sum(CASE WHEN p.Age >= 10 AND p.Age < 20 THEN c.ContributionAmount END) AS [cm],
  Sum(CASE WHEN p.Age >= 20 AND p.Age < 30 THEN c.ContributionAmount END) AS [cn],
  Sum(CASE WHEN p.Age >= 30 AND p.Age < 40 THEN c.ContributionAmount END) AS [co],
  Sum(CASE WHEN p.Age >= 40 AND p.Age < 50 THEN c.ContributionAmount END) AS [cp],
  Sum(CASE WHEN p.Age >= 50 AND p.Age < 60 THEN c.ContributionAmount END) AS [cq],
  Sum(CASE WHEN p.Age >= 60 AND p.Age < 70 THEN c.ContributionAmount END) AS [cr],
  Sum(CASE WHEN p.Age >= 70 AND p.Age < 80 THEN c.ContributionAmount END) AS [cs],
  Sum(CASE WHEN p.Age >= 80 AND p.Age < 90 THEN c.ContributionAmount END) AS [ct],
  Sum(CASE WHEN p.Age >= 90 THEN c.ContributionAmount END) AS [cu],
  Count(Distinct CASE WHEN p.Age >= 0 AND p.Age < 10 THEN p.PeopleID END) AS [pl],
  Count(Distinct CASE WHEN p.Age >= 10 AND p.Age < 20 THEN p.PeopleID END) AS [pm],
  Count(Distinct CASE WHEN p.Age >= 20 AND p.Age < 30 THEN p.PeopleID END) AS [pn],
  Count(Distinct CASE WHEN p.Age >= 30 AND p.Age < 40 THEN p.PeopleID END) AS [po],
  Count(Distinct CASE WHEN p.Age >= 40 AND p.Age < 50 THEN p.PeopleID END) AS [pp],
  Count(Distinct CASE WHEN p.Age >= 50 AND p.Age < 60 THEN p.PeopleID END) AS [pq],
  Count(Distinct CASE WHEN p.Age >= 60 AND p.Age < 70 THEN p.PeopleID END) AS [pr],
  Count(Distinct CASE WHEN p.Age >= 70 AND p.Age < 80 THEN p.PeopleID END) AS [ps],
  Count(Distinct CASE WHEN p.Age >= 80 AND p.Age < 90 THEN p.PeopleID END) AS [pt],
  Count(Distinct CASE WHEN p.Age >= 90 THEN p.PeopleID END) AS [pu]
FROM Contribution c
INNER JOIN
  People p
ON c.PeopleID =  p.PeopleID
--WHERE DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(Year(c.ContributionDate))) >= (DATEPART(Year,getdate())-10)
-- Group By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(Year(c.ContributionDate)))
-- Order By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(Year(c.ContributionDate)))
'''

sqlContributionNumbers = '''
Select
  DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate) AS [UniqueContributorsbyYear],
  Sum(c.ContributionAmount) AS [Contributed],
  Round(Sum(c.ContributionAmount)/Count(Distinct c.ContributionID),2) AS [AverageGift],
  Count(c.ContributionID) AS [Contributions], 
  Count(Distinct c.PeopleID) AS [UniqueGivers],
  Count(c.ContributionID)/Count(Distinct c.PeopleID) AS [AvgNumofGifts]
FROM Contribution c
INNER JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) >= (DATEPART(Year,getdate())-10)
 Group By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate)))
 Order By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate)))
'''

template = '''
<style>
    table.table.no-border td,
    table.table.no-border th {
        border-top: none;
    }
</style>
<table class="table centered">
    <thead>
        {{#with header}}
        <tr>
            <td colspan="7">
                <div class="text-center">
                    <h2>Unique Contributors by Year</h2>
                </div>
                <br>
            </td>
        </tr>
        {{/with}}
        <tr>
            <th>Year</th>
            <th>0-9</th>
            <th>10-19</th>
            <th>20-29</th>
            <th>30-39</th>
            <th>40-49</th>
            <th>50-59</th>
            <th>60-69</th>
            <th>70-79</th>
            <th>80-89</th>
            <th>90+</th>
        </tr>
    </thead>
    <tbody>
    {{#each gifts}}
        <tr style="{{Bold}}">
            <td>{{UniqueContributorsbyYear}}</td>
            <td>{{Fmt ca "N0"}}<br>{{Fmt pa "N0"}}</td>
            <td>{{Fmt cb "N0"}}<br>{{Fmt pb "N0"}}</td>
            <td>{{Fmt cc "N0"}}<br>{{Fmt pc "N0"}}</td>
            <td>{{Fmt cd "N0"}}<br>{{Fmt pd "N0"}}</td>
            <td>{{Fmt ce "N0"}}<br>{{Fmt pe "N0"}}</td>
            <td>{{Fmt cf "N0"}}<br>{{Fmt pf "N0"}}</td>
            <td>{{Fmt cg "N0"}}<br>{{Fmt pg "N0"}}</td>
            <td>{{Fmt ch "N0"}}<br>{{Fmt ph "N0"}}</td>
            <td>{{Fmt ci "N0"}}<br>{{Fmt pi "N0"}}</td>
            <td>{{Fmt cj "N0"}}<br>{{Fmt pj "N0"}}</td>
            <td><b>{{Fmt ck "N0"}}<br>{{Fmt pk "N0"}}</b></td>
       </tr>
    {{/each}}
    {{#each giftstotal}}
        <tr style="{{Bold}}">
            <td><b>Total</b></td>
            <td><b>{{Fmt cl "N0"}}<br>{{Fmt pl "N0"}}</b></td>
            <td><b>{{Fmt cm "N0"}}<br>{{Fmt pm "N0"}}</b></td>
            <td><b>{{Fmt cn "N0"}}<br>{{Fmt pn "N0"}}</b></td>
            <td><b>{{Fmt co "N0"}}<br>{{Fmt po "N0"}}</b></td>
            <td><b>{{Fmt cp "N0"}}<br>{{Fmt pp "N0"}}</b></td>
            <td><b>{{Fmt cq "N0"}}<br>{{Fmt pq "N0"}}</b></td>
            <td><b>{{Fmt cr "N0"}}<br>{{Fmt pr "N0"}}</b></td>
            <td><b>{{Fmt cs "N0"}}<br>{{Fmt ps "N0"}}</b></td>
            <td><b>{{Fmt ct "N0"}}<br>{{Fmt pt "N0"}}</b></td>
            <td><b>{{Fmt cu "N0"}}<br>{{Fmt pu "N0"}}</b></td>
            <td><b></b></td>
       </tr>
    {{/each}}
    </tbody>
</table>
<br>

'''
Data.gifts = q.QuerySql(sqlgifts)
Data.giftstotal = q.QuerySql(sqlgiftstotal)
#Data.contributionnumbers = q.QuerySql(sqlContributionNumbers)
print model.RenderTemplate(template)
