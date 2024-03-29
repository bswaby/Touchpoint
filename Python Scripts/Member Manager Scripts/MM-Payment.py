model.Title = "MM-Payment"
ProgramID = model.Data.ProgramID

for a in q.QuerySql("Select Name From Program Where Id = " + ProgramID):
    ProgramName = a.Name 
    
model.Header = ProgramName + ' | Program Manager'


#make payment
if model.Data.pid == "":
    print ""
elif model.Data.PayAmount == "":
    print '<h2>missing payment</h2>'
elif model.Data.PaymentType == "":
    print '<h2>missing payment type</h2>'
elif model.Data.PaymentDescription == "":
    print '<h2>missing payment description</h2>'
elif model.Data.PaymentOrg == "":
    print '<h2>organization missing</h2>'
else:
    messageDescription = model.Data.PaymentType + model.Data.PaymentDescription
    model.AddTransaction(int(model.Data.pid), int(model.Data.PaymentOrg), float(model.Data.PayAmount), messageDescription)
    #model.Email(3134,3134, "bswaby@fbchtn.org", "Ben Swaby - FBCHville", "Test", messageDescription)
    print '<p><h2>A {0} payment of {1} has been made with (<i>{2}</i>) description</h2></p>'.format(model.Data.PaymentType,model.Data.PayAmount,model.Data.PaymentDescription)
    

print '''<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
        </br></br></br>
        <a href="''' + model.CmsHost + '''/PyScript/MM-MemberManager?ProgramID=''' + ProgramID + '''"><i class="fa fa-home fa-3x"></i></a>'''.format(model.CmsHost,model.Data.pid)