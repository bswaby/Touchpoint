model.Title = "MM-ChargeIndividual"
ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

#make payment
if model.Data.pid == "":
    print ""
elif model.Data.PayAmount == "":
    print '<h2>missing payment</h2>'
elif model.Data.PaymentType == "":
    print '<h2>missing payment type</h2>'
elif model.Data.PaymentDescription == "":
    print '<h2>missing payment description</h2>'
elif model.Data.PaymentOrg == "" or model.Data.PaymentOrg == None:
    print '<h2>organization missing</h2>'
else:
    messageDescription = model.Data.PaymentType + model.Data.PaymentDescription
    model.AdjustFee(int(model.Data.pid), int(model.Data.PaymentOrg), -float(model.Data.PayAmount), messageDescription)
    #model.Email(3134,3134, "bswaby@fbchtn.org", "Ben Swaby - FBCHville", "Test", messageDescription)

    #TODO rewrite message displayed
    print '<p><h2>A charge of ${1} has been made </h2></p>'.format(model.Data.PaymentType, model.Data.PayAmount,model.Data.PaymentDescription)
    
#TODO redirect back to charge
print '''<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
        </br></br></br>
        <a href="''' + model.CmsHost + '''/PyScript/MM-Charge?ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''"><i class="fa fa-home fa-3x"></i></a>'''.format(model.CmsHost,model.Data.pid)
