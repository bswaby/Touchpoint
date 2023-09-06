#Automation Acccount Where Global Values are Stored
AutomationAccount = 40678

ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName
FamilyId = model.Data.FamilyId

if model.Data.FamilyTotals != "":
    FamilyTotals = list((model.Data.FamilyTotals).replace('[', '').replace(']', '').replace('', '').split(','))
else:
    FamilyTotals = []

if model.Data.FamilyOrder != "":
    FamilyOrder = list((model.Data.FamilyOrder).replace('[', '').replace(']', '').replace('', '').split(','))
else:
    FamilyOrder = []

#Get Email from and EmailAddress
EmailFrom = int(model.ExtraValueInt(AutomationAccount, str(ProgramID) + '_EmailFrom'))
getEmailAddress = q.QuerySql("SELECT TOP 1 EmailAddress, FirstName from People Where PeopleId = " + str(EmailFrom))
for em in getEmailAddress:
    Data.email = em.EmailAddress
    
getProgramName = q.QuerySql("Select Name from Program Where Program.Id = " + str(ProgramID))
for pg in getProgramName:
    Data.ProgramName = pg.Name
        
#Build PayLink
paylink = " "
#due = float(model.Data.totaldue)
due = '${:,.2f}'.format(float(model.Data.totaldue))
#if a.OrganizationId != None and a.PeopleId != None:
paylink = model.GetPayLink(int(model.Data.pid), int(model.Data.oid))
paylink = model.GetAuthenticatedUrl(int(model.Data.oid), paylink, True)

sendGroup = q.QuerySqlInt("SELECT TOP 1 ID from SmsGroups")

#build message
message = '''You have open balance of due with the <b>''' + Data.ProgramName + '''</b> minsitry at FBCHville.
 Click the following link to pay the outstanding balance of {0}.  {1}'''.format(due,paylink) 
messagesms = '''You an open balance with the ''' + Data.ProgramName + ''' at FBCHville.  Click the following link to pay the outstanding balance of {0}.  {1}'''.format(due,paylink) 
#print(int(model.Data.pid), EmailFrom, Data.email, "need to automate program name - FBCHville", "FBCHville Open Invoice", message)

#get specific billing or head of household billing
AltPayID = int(model.Data.AltPayID)

if AltPayID == 0: #head of household
    HoHPeople = q.QuerySql("Select PeopleId, EmailAddress, FirstName, LastName, CellPhone from People Where FamilyId = "+ FamilyId + " and PositionInFamilyId = 10")
    for hoh in HoHPeople:
        Data.PeopleId = hoh.PeopleId
        Data.EmailAddress = hoh.EmailAddress
        Data.FirstName = hoh.FirstName
        Data.LastName = hoh.LastName
        Data.CellPhone = hoh.CellPhone

else: #Specific Pay
    HoHPeople = q.QuerySql("Select PeopleId, EmailAddress, FirstName, LastName, CellPhone from People Where PeopleId = "+ str(AltPayID))
    for hoh in HoHPeople:
        Data.PeopleId = hoh.PeopleId
        Data.EmailAddress = hoh.EmailAddress
        Data.FirstName = hoh.FirstName
        Data.LastName = hoh.LastName
        Data.CellPhone = hoh.CellPhone

if len(FamilyTotals) > 0:
    for i in range(len(FamilyOrder)):
        if FamilyOrder[i] != Data.PeopleId:
            model.AdjustFee(int(FamilyOrder[i]), int(model.Data.oid), float(FamilyTotals[i]), "move-to-payer charge")
            model.AdjustFee(int(Data.PeopleId), model.ExtraValueIntOrg(int(model.Data.oid), 'payerInvolvement'), -float(FamilyTotals[i]), "move-to-payer charge")
else:
    totDue = float(model.Data.totaldue)
    print(model.ExtraValueIntOrg(int(model.Data.oid), 'payerInvolvement'))
    print(model.Data.pid)
    print(Data.PeopleId)

    #TODO need to reconsider this logic since the payer is outside of the involvement
    model.AdjustFee(int(model.Data.pid), int(model.Data.oid), totDue, "move-to-payer charge")
    model.AdjustFee(int(Data.PeopleId), model.ExtraValueIntOrg(int(model.Data.oid), 'payerInvolvement'), -totDue, "move-to-payer charge")

if Data.EmailAddress != None:
    model.Email(int(Data.PeopleId), EmailFrom, Data.email, Data.ProgramName + " - FBCHville", "FBCHville Open Invoice", message)
    print ('''<h2>Paylink sent to {0} {1}'s email address of {2}.<h2>''').format(Data.FirstName,Data.LastName,Data.EmailAddress)
else: 
    print ('''<h2>Account does not have an email address associated with it</h2>''')

if Data.CellPhone != None:
    model.SendSms(int(Data.PeopleId), sendGroup, "Open Invoice", messagesms)
    print ('''<h2>Paylink sent to {0} {1}'s cell phone number ({2}).<h2>''').format(Data.FirstName,Data.LastName,Data.CellPhone)
else:
    print ('''<h2>Account does not have an cell phone associated with it</h2>''')

#print '''</br></br><link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">'''
print ('''<a href="''' + model.CmsHost + '''/PyScript/MM-MemberManager?ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''"><i class="fa fa-home fa-3x"></i></a>'''.format(model.CmsHost,model.Data.pid))
