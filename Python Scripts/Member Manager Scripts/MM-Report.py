ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

model.Header = ProgramName +' Reports'

template = '''
<a onclick="history.back()"><i class="fa fa-hand-o-left fa-2x" aria-hidden="true"></i></a>&nbsp;&nbsp;
<a href="https://myfbch.com/PyScript/MM-MemberManager?ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''"><i class="fa fa-home fa-2x"></i></a>
<h4>Activity Log</h2>
<a href="https://myfbch.com/CheckinTimes" target="_blank">Activity</a></br>
<h4>Week over Week</h2>
<a href="https://myfbch.com/RunScript/Report%20-%20WoW%20By%20Age%20Bin/" target="_blank">Age</a></br>
<a href="https://myfbch.com/RunScript/Report%20-%20WoW%20Hourly/" target="_blank">Hourly Bins</a></br>
<a href="https://myfbch.com/RunScript/Report%20-%20Day%20of%20the%20Week%20Attendance/" target="_blank">Day of the Week</a></br>
<h4>Month over Month</h2>
<a href="https://myfbch.com/RunScript/Report%20-%20MoM%20by%20Age/" target="_blank">Age</a></br>
'''

SReport = model.RenderTemplate(template)
print SReport