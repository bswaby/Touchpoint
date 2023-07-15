model.JoinOrg(int(model.Data.addorg),int(model.Data.pid))
ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

print '''
<h3>Added to Organization</h3></br></br>
    </br></br><link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">  
    <button onclick="window.location.href='{0}/PyScript/MM-MemberDetails?p1={1}&FamilyId={2}&ProgramName={3}&ProgramID={4}';"><i class="fa fa-hand-o-left" aria-hidden="true"></i></button>
    <button onclick="window.location.href='{0}/PyScript/MM-MemberManager?&ProgramName={3}&ProgramID={4}';"><i class="fa fa-home"></i></button>'''.format(model.CmsHost,model.Data.pid, model.Data.FamilyId, ProgramName, ProgramID)