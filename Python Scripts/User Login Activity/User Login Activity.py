model.Header = 'User Login Activity'

sqlfailedloginstat = '''
Select 
    Count(Activity) AS [FailedLogins],
    Activity,
    Max(ActivityDate) AS [LastFailedAttempt]
from dbo.activitylog 
Where 
    (Activity Like '%Failed password%' 
        OR Activity Like '%Invalid log%'
        OR Activity Like '%ForgotPassword%')
    AND CAST(ActivityDate as Date) >= CAST(DATEADD(DAY, -3, CONVERT(DATE, GETDATE())) AS Date) --CAST(GETDATE() as Date)
Group By Activity
Order by Max(ActivityDate) Desc;
'''

sqlfailedlogins = '''
Select 
	Top 200 * 
From 
	dbo.activitylog 
Where 
	Activity Like '%ForgotPassword%'
	OR Activity Like '%failed password%'
	OR Activity Like '%Invalid log%'
Order by activitydate Desc
	---OR Activity Like '%OnlineReg Login%'
'''

sqllogins = '''
Select 
	Top 50 * 
From 
	dbo.activitylog 
Where 
	Activity Like '%logged in%'
Order by activitydate Desc
'''


    
template = '''
<style>
    div.scroll {
        bckground-color: #fed9ff;
        width: 850px;
        height: 450px;
        overflow-x: hidden;
        overflow-y: auto;
        text-align: center;
        padding: 5px;
    }
</style>
<h4>Misc Reports</h4>
<a href="https://myfbch.com/RunScript/Tech_UserStat/" target="_blank">User Account Info</a>
<br>
<h4>Failed Logins</h4>
<h6>--Last 72hr Login Failures--</h6>
<table width="500px"; border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top;>Failed Times</td>
        <td border: 2px solid #000; vertical-align:top;>Failure</td>
        <td width="160px"; border: 2px solid #000; vertical-align:top;>Last Failed Attempt</td>
    </tr>
    {{#each failedloginstat}}
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top;>{{FailedLogins}}</td>
        <td border: 2px solid #000; vertical-align:top;>{{Activity}}</td>
        <td width="160px"; border: 2px solid #000; vertical-align:top;>{{LastFailedAttempt}}</td>
    </tr>
    {{/each}}
    </tbody>
</table>
<br>
<h6>--Last 200 failed attempt logs--</h6>
<div class="scroll">
<table width="800px"; height="500px"; border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    <tr style="{{Bold}}">
        <td width="160px"; border: 2px solid #000; vertical-align:top;><h6>Date</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>User Id</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>Activity</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>People Id</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>Involvement</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>Ip Address</h6></td>
    </tr>
    {{#each failedlogins}}
    <tr style="{{Bold}}">
        <td width="160px"; border: 2px solid #000; vertical-align:top;>{{ActivityDate}}</td>
        <td  border: 2px solid #000; vertical-align:top;>{{UserId}}</td>
        <td  border: 2px solid #000; vertical-align:top;>{{Activity}}</td>
        <td  border: 2px solid #000; vertical-align:top;><a href="https://myfbch.com/Person2/{{PeopleId}}" target="_blank">{{PeopleId}}</a></td>
        <td  border: 2px solid #000; vertical-align:top;><a href="https://myfbch.com/Org/{{OrgId}}" target="_blank">{{OrgId}}</a></td>
        <td  border: 2px solid #000; vertical-align:top;><a href="https://mxtoolbox.com/SuperTool.aspx?action=ptr%3a{{ClientIp}}&run=toolpage" target="_blank">{{ClientIp}}</a></td>
    </tr>
    {{/each}}
    </tbody>
</table>
</div>
<br>
<h4>Last 50 Logins</h4>
<div class="scroll">
<table width="600px"; border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    <tr style="{{Bold}}">
        <td width="160px"; border: 2px solid #000; vertical-align:top;><h6>Date</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>User Id</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>Activity</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>People Id</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>Involvement</h6></td>
        <td  border: 2px solid #000; vertical-align:top;><h6>Ip Address</h6></td>
    </tr>
    {{#each logins}}
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top;>{{ActivityDate}}</td>
        <td  border: 2px solid #000; vertical-align:top;>{{UserId}}</td>
        <td  border: 2px solid #000; vertical-align:top;>{{Activity}}</td>
        <td  border: 2px solid #000; vertical-align:top;><a href="https://myfbch.com/Person2/{{PeopleId}}" target="_blank">{{PeopleId}}</a></td>
        <td  border: 2px solid #000; vertical-align:top;><a href="https://myfbch.com/Org/{{OrgId}}" target="_blank">{{OrgId}}</a></td>
        <td  border: 2px solid #000; vertical-align:top;><a href="https://mxtoolbox.com/SuperTool.aspx?action=ptr%3a{{ClientIp}}&run=toolpage" target="_blank">{{ClientIp}}</a></td>
    </tr>
    {{/each}}
    </tbody>
</table>
</div>


'''


Data.failedloginstat = q.QuerySql(sqlfailedloginstat)
Data.failedlogins = q.QuerySql(sqlfailedlogins)
Data.logins = q.QuerySql(sqllogins)
SReport = model.RenderTemplate(template)
print(SReport)

#
#NMReport = model.RenderTemplate(template)
#print(NMReport)
