ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

model.Header = ProgramName + ' Fee "The Terminator"'

sqlmemberstat = '''
SELECT 
    SUM( 
    CASE 
        WHEN DATEPART(YEAR, dbo.People.JoinDate) = DATEPART(YEAR, GETDATE())
        AND DATEPART(MONTH, dbo.People.JoinDate) = DATEPART(MONTH, GETDATE())
        THEN 1 
        ELSE 0 
    END) AS [CurrentMonth], 
    SUM( 
    CASE 
        WHEN dbo.People.JoinDate BETWEEN '20221001' AND '20230930' 
        THEN 1 
        ELSE 0 
    END) AS [FiscalYear23], 
    SUM( 
    CASE 
        WHEN dbo.People.JoinDate BETWEEN '20211001' AND '20220930' 
        THEN 1 
        ELSE 0 
    END) AS [FiscalYear22],
    SUM( 
    CASE 
        WHEN dbo.People.JoinDate BETWEEN '20201001' AND '20210930' 
        THEN 1 
        ELSE 0 
    END) AS [FiscalYear21],  
    SUM( 
    CASE 
        WHEN dbo.People.JoinDate BETWEEN '20191001' AND '20200930' 
        THEN 1 
        ELSE 0 
    END) AS [FiscalYear20],
    SUM(CASE 
        WHEN dbo.People.MemberStatusId = 10 
        THEN 1 
        ELSE 0 
    END) AS [TotalMembers] 
FROM 
    dbo.People;
'''

sqlmembers = '''
SELECT 
    COUNT(dbo.People.IsDeceased), 
    dbo.People.PeopleId, 
    dbo.People.Name, 
    dbo.People.LastName,
    dbo.People.FirstName,
    dbo.People.Age, 
    dbo.People.FamilyId,
    dbo.People.PrimaryAddress, 
    dbo.People.PrimaryCity, 
    dbo.People.PrimaryState, 
    dbo.People.PrimaryZip, 
    dbo.FmtPhone(dbo.People.HomePhone) AS [HomePhone], 
    dbo.FmtPhone(dbo.People.CellPhone) AS [CellPhone], 
    dbo.People.EmailAddress, 
    dbo.People.JoinDate, 
    lookup.JoinType.Description                                AS [HowJoined], 
    STRING_AGG(Organizations_alias1.OrganizationName,' | ') AS [orgs1]
FROM 
    dbo.OrganizationMembers 
RIGHT JOIN 
    dbo.People 
ON 
    ( 
        dbo.OrganizationMembers.PeopleId = dbo.People.PeopleId) 
LEFT JOIN 
    dbo.Organizations Organizations_alias1 
ON 
    ( 
        dbo.OrganizationMembers.OrganizationId = Organizations_alias1.OrganizationId) 
INNER JOIN 
    lookup.JoinType 
ON 
    ( 
        dbo.People.JoinCodeId = lookup.JoinType.Id) 
WHERE 
    dbo.People.JoinDate BETWEEN DATEADD(DAY, -7, CONVERT(DATE, GETDATE())) AND CONVERT(DATE, 
    GETDATE()) 
GROUP BY 
    dbo.People.LastName,
    dbo.People.FirstName,
    dbo.People.Name, 
    dbo.People.PeopleId, 
    dbo.People.Age, 
    dbo.People.FamilyId,
    dbo.People.PrimaryAddress, 
    dbo.People.PrimaryCity, 
    dbo.People.PrimaryState, 
    dbo.People.PrimaryZip, 
    dbo.People.HomePhone, 
    dbo.People.CellPhone, 
    dbo.People.EmailAddress, 
    dbo.People.JoinDate, 
    lookup.JoinType.Description
Order By dbo.People.LastName, dbo.People.FirstName;
'''
    
template = '''
<table class="table centered"; border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    <tr>
        <td><h4>Name</h4></td>
        <td><h4>Address</h4></td>
        <td><h4>Communication</h4></td>
        <td><h4>Involvements</h4></td>
    </tr>
    {{#each gifts}}
    <tr style="{{Bold}}">
            <td  border: 2px solid #000; vertical-align:top; style="background-color: #eeeeee"><a href=https://myfbch.com/Person2/{{PeopleId}}><font size=5>{{Name}}</font></a> - 
                <br /><font size=3>{{HowJoined}}
                <br />Age:{{Age}}
                <br />Family:{{FamilyId}}</td></font>
            <td border: 2px solid #000; vertical-align:top;><a href="https://www.google.com/maps/search/{{PrimaryAddress}},{{PrimaryCity}},{{PrimaryState}},{{PrimaryZip}}"><font size=3>{{PrimaryAddress}}</a>
                <br />{{PrimaryCity}}, {{PrimaryState}}, {{PrimaryZip}}</td></font>
            <td border: 2px solid #000; vertical-align:top;>h:<a href=tel:{{HomePhone}}><font size=3>{{HomePhone}}</a>
                <br />c:<a href=tel:{{CellPhone}}>{{CellPhone}}</a>
                <br /><a href=mailto:{{EmailAddress}}>{{EmailAddress}}</a></td></font>
            <td border: 2px solid #000; vertical-align:top;>{{orgs1}}</td>
    </tr>
   {{/each}}
    </tbody>
</table>
'''

mstattemplate = '''
<form><input type="button" value=" < " onclick="history.back()">
<button onclick="window.location.href='https://myfbch.com/PyScript/MM-MemberManager?ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''';"><i class="fa fa-home"></i></button></form>
<h2>New Member Report</h2>
<table ; width="225px" border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
{{#each stat}}
    <tr style="{{Bold}}">
            <tr><td><h5>Period</h5></td>
            <td style="text-align:center;">Total Members</td></tr>
            <tr><td style="background-color: #eeeeee"><h5><b>Current Month</b></h5></td>
            <td border: 2px solid #000; vertical-align:top; width="75px" style="text-align:center;">{{CurrentMonth}}</td></tr>
            <tr><td style="background-color: #eeeeee"><h5><b>2023 Fiscal Year</b></h5></td>
            <td border: 2px solid #000; vertical-align:top; width="75px" style="text-align:center;">{{FiscalYear23}}</td></tr>
            <tr><td style="background-color: #eeeeee"><h5><b>2022 Fiscal Year</b></h5></td>
            <td border: 2px solid #000; vertical-align:top; width="75px" style="text-align:center;">{{FiscalYear22}}</td></tr>
            <tr><td style="background-color: #eeeeee"><h5><b>2021 Fiscal Year</b></h5></td>
            <td border: 2px solid #000; vertical-align:top; width="75px" style="text-align:center;">{{FiscalYear21}}</td></tr>
            <tr><td style="background-color: #eeeeee"><h5><b>2020 Fiscal Year</b></h5></td>
            <td border: 2px solid #000; vertical-align:top; width="75px" style="text-align:center;">{{FiscalYear20}}</td></tr>
            <tr><td style="background-color: #eeeeee"><h5><b>Total Members</b></h5></td>
            <td border: 2px solid #000; vertical-align:top;width="75px" style="text-align:center;">{{TotalMembers}}</td></tr>
   </tr>
{{/each}}
    </tbody>
</table>
<br />
<br />
'''

Data.stat = q.QuerySql(sqlmemberstat)
SReport = model.RenderTemplate(mstattemplate)
#print(SReport)

Data.gifts = q.QuerySql(sqlmembers)
NMReport = model.RenderTemplate(template)
#print(NMReport)

NewMemberReport = SReport + NMReport

#model.Email("NewMemberTeam", 3134, "chodges@fbchtn.org", "Cindy Hodges", "New Members - Last 7 Days", NewMemberReport)
#print(Data.gifts)