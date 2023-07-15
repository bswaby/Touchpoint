
ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

paylink = " "
#due = float(model.Data.totaldue)
due = '${:,.2f}'.format(float(model.Data.totaldue))
#if a.OrganizationId != None and a.PeopleId != None:
paylink = model.GetPayLink(int(model.Data.pid), int(model.Data.oid))
paylink = model.GetAuthenticatedUrl(int(model.Data.oid), paylink, True)

sendGroup = q.QuerySqlInt("SELECT TOP 1 ID from SmsGroups")

message = '''Click to pay the outstanding balance of {0}.  {1}'''.format(due,paylink) 

if model.Data.sendemail == "y":

    if ProgramID == 1108:
      model.Email(a.PeopleId, 3414, "dmeyer@fbchtn.org", "David Meyer - FBCHville", "FBCHville Receipt", message)
    elif ProgramID == 1109:
      model.Email(a.PeopleId, 7365, "tklapwyk@fbchtn.org", "Tammy Klapwyk - FBCHville", "FBCHville Receipt", message)
    elif ProgramID == 1143:
      model.Email(a.PeopleId, 11180, "lpoteet@fbchtn.org", "Laura Poteet - FBCHville", "FBCHville Receipt", message)
    elif ProgramID == 1149:
      model.Email(a.PeopleId, 36153, "sgilmore@fbchtn.org", "Shannon Gilmore - FBCHville", "FBCHville Receipt", message)
    elif ProgramID == 1152:
      model.Email(a.PeopleId, 14221, "tbeals@fbchtn.org", "Tucker Beals - FBCHville", "FBCHville Test Receipt", message)

    print '''<h2>Paylink link to email<h2></br>'''
else: 
    print '''<h2>Account does not have an email address associated with it</h2>'''

if model.Data.sendtext == "y":
    model.SendSms(int(model.Data.pid), sendGroup, "Open Invoice", message)
    print '''<h2>Paylink link to text<h2></br>'''
else:
    print '''<h2>Account does not have an cell phone associated with it</h2>'''


print '''</br></br><link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">'''    
print '''<input type="button" value=" < " onclick="history.back()">
<button onclick="window.location.href='{0}/PyScript/MM-MemberManager?ProgramName={1}&ProgramID={2}';"><i class="fa fa-home"></i></button>'''.format(model.CmsHost, ProgramName, ProgramID)