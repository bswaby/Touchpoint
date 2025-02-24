#Roles=rpt_BusinessAdmin

#Update fiscal month to your need.  If you want calendar year set to 0.  If you want August as your start month, count backwards..
#so in the case of an August fiscal year, you would put in 5.

model.Header = 'Contribution By Age Bin with Distinct Givers'

fiscalmonth = '3'
NonContribution = '99'  #this should be the default to exclude non-contributions
GeneralFundId = '1' #add comma seperated funds if you want to limit.  For example, we only want to show our general fund

sqlgifts = """
SELECT  
  DATEPART(Year, DATEADD(month, {0}, c.ContributionDate)) AS [UniqueContributorsbyYear],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 0 AND 9 THEN c.ContributionAmount END) AS [ca],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 10 AND 19 THEN c.ContributionAmount END) AS [cb],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 20 AND 29 THEN c.ContributionAmount END) AS [cc],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 30 AND 39 THEN c.ContributionAmount END) AS [cd],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 40 AND 49 THEN c.ContributionAmount END) AS [ce],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 50 AND 59 THEN c.ContributionAmount END) AS [cf],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 60 AND 69 THEN c.ContributionAmount END) AS [cg],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 70 AND 79 THEN c.ContributionAmount END) AS [ch],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 80 AND 89 THEN c.ContributionAmount END) AS [ci],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       >= 90 THEN c.ContributionAmount END) AS [cj],

    SUM(c.ContributionAmount) AS [ck],

    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 0 AND 9 THEN p.PeopleID END) AS [pa],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 10 AND 19 THEN p.PeopleID END) AS [pb],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 20 AND 29 THEN p.PeopleID END) AS [pc],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 30 AND 39 THEN p.PeopleID END) AS [pd],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 40 AND 49 THEN p.PeopleID END) AS [pe],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 50 AND 59 THEN p.PeopleID END) AS [pf],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 60 AND 69 THEN p.PeopleID END) AS [pg],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 70 AND 79 THEN p.PeopleID END) AS [ph],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 80 AND 89 THEN p.PeopleID END) AS [pi],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         >= 90 THEN p.PeopleID END) AS [pj],

  COUNT(DISTINCT p.PeopleID) AS [pk]

FROM Contribution c
INNER JOIN People p ON c.PeopleID = p.PeopleID
WHERE DATEPART(Year, DATEADD(month, {0}, c.ContributionDate)) >= (DATEPART(Year, GETDATE()) - 10)
    {1}
    AND c.ContributionTypeId <> {2}
GROUP BY DATEPART(Year, DATEADD(month, {0}, c.ContributionDate))
ORDER BY DATEPART(Year, DATEADD(month, {0}, c.ContributionDate));
"""

sqlgiftstotal = """
SELECT  
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 0 AND 9 THEN c.ContributionAmount END) AS [cl],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 10 AND 19 THEN c.ContributionAmount END) AS [cm],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 20 AND 29 THEN c.ContributionAmount END) AS [cn],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 30 AND 39 THEN c.ContributionAmount END) AS [co],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 40 AND 49 THEN c.ContributionAmount END) AS [cp],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 50 AND 59 THEN c.ContributionAmount END) AS [cq],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 60 AND 69 THEN c.ContributionAmount END) AS [cr],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 70 AND 79 THEN c.ContributionAmount END) AS [cs],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       BETWEEN 80 AND 89 THEN c.ContributionAmount END) AS [ct],
    SUM(CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
       - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                           DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
       >= 90 THEN c.ContributionAmount END) AS [cu],

    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) < 10 
        THEN p.PeopleID END) AS [pl],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 10 AND 19 THEN p.PeopleID END) AS [pm],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 20 AND 29 THEN p.PeopleID END) AS [pn],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 30 AND 39 THEN p.PeopleID END) AS [po],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 40 AND 49 THEN p.PeopleID END) AS [pp],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 50 AND 59 THEN p.PeopleID END) AS [pq],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 60 AND 69 THEN p.PeopleID END) AS [pr],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 70 AND 79 THEN p.PeopleID END) AS [ps],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         BETWEEN 80 AND 89 THEN p.PeopleID END) AS [pt],
    COUNT(DISTINCT CASE WHEN (DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate) 
         - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), c.ContributionDate), 
                             DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay)) > c.ContributionDate THEN 1 ELSE 0 END) 
         >= 90 THEN p.PeopleID END) AS [pu]
FROM Contribution c
INNER JOIN People p ON c.PeopleID = p.PeopleID
Where c.ContributionTypeId <> {1}
 {0}
"""

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

if GeneralFundId:
    GeneralFundId = """AND c.FundId IN ({0})""".format(GeneralFundId)
else:
    GeneralFundId = ''
    
Data.gifts = q.QuerySql(sqlgifts.format(fiscalmonth,GeneralFundId,NonContribution))
Data.giftstotal = q.QuerySql(sqlgiftstotal.format(GeneralFundId,NonContribution))
print model.RenderTemplate(template)
