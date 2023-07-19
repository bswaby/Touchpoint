ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

model.Header = ProgramName + ' Member Manager'

listsql = '''
DECLARE @ProgramID int = ''' + ProgramID +'''
SELECT 
    dbo.People.PeopleId,
    dbo.People.Name, 
    dbo.People.FamilyId, 
    dbo.People.FirstName, 
    dbo.People.LastName, 
    dbo.People.Age,
    dbo.People.EmailAddress,
    dbo.People.PrimaryAddress,
    dbo.People.PrimaryCity,
    dbo.People.PrimaryState,
    dbo.People.PrimaryZip,
    dbo.People.CellPhone,
    dbo.People.HomePhone,
    Organizations_alias1.OrganizationName,
    Organizations_alias1.OrganizationId,
    dbo.MemberTags.Name AS [SubGroup],
	dbo.TransactionSummary.TotDue
FROM 
    dbo.ProgDiv 
INNER JOIN dbo.Program ON (dbo.ProgDiv.ProgId = dbo.Program.Id) 
INNER JOIN dbo.Division ON (dbo.ProgDiv.DivId = dbo.Division.Id) 
INNER JOIN dbo.Organizations Organizations_alias1 ON (dbo.Division.Id = Organizations_alias1.DivisionId) 
INNER JOIN dbo.OrganizationMembers ON (Organizations_alias1.OrganizationId = dbo.OrganizationMembers.OrganizationId) 
LEFT JOIN dbo.TransactionSummary ON (dbo.OrganizationMembers.TranId = dbo.TransactionSummary.RegId) 
INNER JOIN dbo.People ON (dbo.OrganizationMembers.PeopleId = dbo.People.PeopleId) 
LEFT JOIN dbo.OrgMemMemTags ON (dbo.OrgMemMemTags.PeopleId = dbo.People.PeopleId) AND (dbo.OrgMemMemTags.OrgId = Organizations_alias1.OrganizationId)
LEFT JOIN dbo.MemberTags ON (dbo.MemberTags.Id = dbo.OrgMemMemTags.MemberTagId)
LEFT JOIN dbo.OrganizationExtra ON Organizations_alias1.OrganizationId = OrganizationExtra.OrganizationID
WHERE 
    dbo.Program.Id = @ProgramID --AND EXISTS (Select CheckInTimes_alias1.CheckInTime From dbo.CheckInTimes Where dbo.People.PeopleId = CheckInTimes_alias1.PeopleId)
    AND OrganizationExtra.Field = 'MemberManagerEnabled' 
    AND OrganizationExtra.BitValue = 1
GROUP BY Organizations_alias1.OrganizationId, 
    Organizations_alias1.OrganizationName, 
    dbo.People.Name, 
    dbo.People.FamilyId, 
    dbo.People.PeopleId, 
    dbo.People.FirstName, 
    dbo.People.LastName, 
    dbo.People.EmailAddress, 
    dbo.People.Age,
    dbo.MemberTags.Name,
	dbo.TransactionSummary.TotDue,
	dbo.People.EmailAddress,
    dbo.People.PrimaryAddress,
    dbo.People.PrimaryCity,
    dbo.People.PrimaryState,
    dbo.People.PrimaryZip,
    dbo.People.CellPhone,
    dbo.People.HomePhone
ORDER BY 
	dbo.TransactionSummary.TotDue Desc,
    dbo.People.LastName ASC, 
    dbo.People.FirstName ASC;
'''


print ('''
<script src ="https://code.jquery.com/jquery-3.5.1.min.js"></script>
<script src ="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"> </script> 
<!-- jQuery Modal -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css" />

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
<a href="'''+ model.CmsHost + '''/PyScript/MM-Balance"><i class="fa fa-usd fa-2x" class=”button-solid”></i></a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="'''+ model.CmsHost + '''/PyScript/MM-Reports"><i class="fa fa-stack-exchange fa-2x" aria-hidden="true"></i></a>
<style>    
button-solid {
  padding: 5px 10px;
  background-color: #ddd;
  color: #000;
}
.button-solid:focus {
    border: none;
    outline: none;
}
h2 {    
  margin-top: 40px;    
  text-transform: none;    
  font-size: 1.75em;    
  font-weight: bold;    
   font-family: "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif;    
  color: white;     
  letter-spacing: -0.005em;     
  word-spacing: 1px;    
  letter-spacing: none;    
  text-align: center;    
}    
h4 {    
  margin-top: 40px;    
  text-transform: none;    
  font-size: 1.75em;    
  font-weight: bold;    
   font-family: "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif;    
  color: #999;     
  letter-spacing: -0.005em;     
  word-spacing: 1px;    
  letter-spacing: none;    
  text-align: center;    
  color: white;    
}    
body {  
  background: #ddd;  
  font-family: sans-serif;  
}  
.container {  
  margin: 20px 0 20px 0;  
}  
.table-filter {  
  margin-bottom: 5px;  
}  
table {  
  box-shadow: 0px 0px 20px rgba(0, 0, 0, 0.1), 0px 10px 20px rgba(0, 0, 0, 0.05),  
    0px 20px 20px rgba(0, 0, 0, 0.05), 0px 30px 20px rgba(0, 0, 0, 0.05);  
}  
th {  
  color: #ffffff;  
  background: #003d4c;  
  font-weight: 700;  
}  
tr {  
  background: #fff;  
}  
tr:hover {  
  background: #f4f4f4;  
}  
th {  
  padding: 5px 5px 5px 10px;  
}  
td {  
  padding: 5px 5px 5px 10px;  
}  
td  {  
  word-wrap: break-word;  
  border-bottom: 1px solid #ccc;  
}  
</style>  
<body>  
<section class = "container">   
  <input  type = "search"  
    class = "form-control table-filter"  
    placeholder = "Search..."  
  />  
  <div class = "table-responsive">  
    <table role = "table" class = "table filtered-table">  
      <thead role = "rowgroup">  
        <tr role = "row">  
          <th role = "columnheader"> Name </th>  
          <th role = "columnheader"> Involvement (subgroup) </th>
          <th role = "columnheader"> Outstanding </th>
          <th role = "columnheader"> Pay By </th>
        </tr>  
      </thead>  
      <tbody role = "rowgroup">  

''')


for a in q.QuerySql(listsql):
    paylink = " "
    AltPayIDNote = " "
    paylinkauth = " "
    sendemail = "n"
    sendtext = "n"
    if a.TotDue != None:
        due = '${:,.2f}'.format(a.TotDue)
        #if a.OrganizationId != None and a.PeopleId != None:
        paylink = model.GetPayLink(a.PeopleId, a.OrganizationId)
        paylinkauth = model.GetAuthenticatedUrl(a.OrganizationId, paylink, True)
        #paylink = a.PeopleId + " : " + a.OrganizationID

    AltPayID = model.ExtraValueInt(int(a.PeopleId), str(ProgramID) + '_AltPayID')
    if AltPayID != 0:
        AltPayPerson = q.QuerySql("Select PeopleId, EmailAddress, FirstName, LastName, CellPhone from People Where PeopleId = "+ str(AltPayID))
        for hoh in AltPayPerson:
            Data.PeopleId = hoh.PeopleId
            Data.EmailAddress = hoh.EmailAddress
            Data.FirstName = hoh.FirstName
            Data.LastName = hoh.LastName
            Data.CellPhone = hoh.CellPhone
            AltPayIDNote = '<i>alternate pay: </i>' + Data.FirstName + ' ' + Data.LastName + '<br>'
    else:
        AltPayIDNote = ""
    
    print '<tr role = "row">'
    print '''<td role = "cell">
      <a href="{12}/PyScript/MM-MemberDetails?p1={1}&FamilyId={3}&ProgramName={4}&ProgramID={5}"> {0} ({2})</a>
      &nbsp<a href="{12}/Person2/{1}#tab-current" target="_blank"><i class="fa fa-info-circle" aria-hidden="true"></i></a>
      <br>{14} email:{6}<br>{7}<br>{8}, {9}, {10}<br>cell:{11}</td>'''.format(a.Name, a.PeopleId, a.Age, a.FamilyId, ProgramName, ProgramID,a.EmailAddress,a.PrimaryAddress,a.PrimaryCity,a.PrimaryState,a.PrimaryZip,a.CellPhone,model.CmsHost, AltPayID, AltPayIDNote)
    print ('<td role = "cell"><a href="') + model.CmsHost + '/Org/{2}" target="_blank">{0}<br>({1})</a></td>'.format(a.OrganizationName,a.SubGroup,a.OrganizationId)
    print ('<td role = "cell">{0}</td>').format(a.TotDue)
    print '<td>'
    if a.EmailAddress != None:
        sendemail = "y"
    
    if a.CellPhone != None:
        sendtext = "y"
    
    if a.TotDue != 0 or a.TotDue != 0.0000:
        if paylinkauth != " ":
            print '<a href="MM-PaymentNotify?pid={0}&totaldue={1}&oid={2}&sendemail={3}&sendtext={4}&ProgramName={5}&ProgramID={6}&AltPayID={7}&FamilyId={8}"><i class="fa fa-credit-card-alt fa-3x" aria-hidden="true"></i></a></br>'.format(a.PeopleId, a.TotDue, a.OrganizationId,sendemail,sendtext,ProgramName,ProgramID,AltPayID,a.FamilyId)
            print '''
              <br>
                <form id="payfee{1}{2}" class="modal" action="MM-Payment">
                  <div class="modalparagraph">
                  <input type="hidden" id="ProgramName" name="ProgramName" value="{3}">
                  <input type="hidden" id="ProgramID" name="ProgramID" value="{4}">
                   <h3>Amount Due:{0}</h3>
                   <input type="hidden" id="pid" name="pid" value="{1}">
                   <input type="hidden" id="PaymentOrg" name="PaymentOrg" value="{2}">
                   <input type="hidden" id="addpayment" name="addpayment" value="y">
                  Payment Type:
                   <input type="radio" name="PaymentType" value="CSH|" id="PaymentType">CASH
                   <input type="radio" name="PaymentType" value="CHK|" id="PaymentType">CHECK
                  </div>
                  <div class="modalparagraph">
                  Description:
                   <input type="text" name="PaymentDescription" id="PaymentDescription"/>
                  </div>
                  <div class="modalparagraph">
                  Pay Amount:<input type="number" name="PayAmount" step="any" id="PayAmount"/>
                  </div>
                  <button>Submit</button>
                </form>
                <a href="#payfee{1}{2}" rel="modal:open"><i class="fa fa-money fa-4x" aria-hidden="true"></i></a></br>'''.format(a.TotDue, a.PeopleId, a.OrganizationId,ProgramName,ProgramID)
    print '</td>'
    print ('<tr>')


print ('''
      </tbody>  
    </table>  
  </div>  
</section>  
<script>
(function ($) {  
  $(document).ready(function () {  
    $('.table-filter').on('input', function () {  
      var value = $(this).val().toLowerCase();  
      $('.filtered-table tbody tr').filter(function () {  
        $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1);  
      });  
    });  
  });  
})(window.jQuery);  
</script>  
''')

if model.HttpMethod == 'post' and model.Data.a == "send":  # Posting a message to be sent
    
    print " "
    # TODO verify that person should be allowed to send such a thing, to the given recipient. 
    
    sendGroup = q.QuerySqlInt("SELECT TOP 1 ID from SmsGroups")  # TODO eventually: make this vaguely intelligent or something
    # For testing, change model.Data.to to model.UserPeopleId to send the message to the active user.
    #link = model.GetPayLink(peopleId, orgId)
    #link = model.GetAuthenticatedUrl(peopleId, link, True)
    
    #model.SendSms(peopleId, sendGroupId, title, "Please pay the balance here: " + link)
    #print model.SendSms('PeopleId = {}'.format(model.Data.to), sendGroup, "Open Invoice", model.Data.message)
    print model.SendSms('PeopleId = {}'.format(model.Data.to), sendGroup, "Open Invoice", model.Data.message)
    print model.Email('PeopleId = {}'.format(model.Data.to),3134, "bswaby@fbchtn.org", "Ben Swaby - FBCHville", "Outstanding " + model.Data.ProgramName + " Fees", model.Data.message)
    
    
    #print('PeopleId = {}'.format(model.Data.to), sendGroup, "Open Invoice", model.Data.message)
    
                
print ('''
    <script defer src="//cdn.jsdelivr.net/npm/sweetalert2@10"></script>
    <script>
    let prepText = function(to, paylnk, due) {
            
        function getPyScriptAddress() {
            let path = window.location.pathname;
            return path.replace("/PyScript/", "/PyScriptForm/");
        }
            
        Swal.fire({
            title: "Payment",
            input: 'textarea',
            inputAttributes: {
                autocapitalize: 'off'
            },
            footer: "Message will be sent from the church's number.  You will not receive replies.",
            inputValue: "Click to pay the outstanding balance of $"  + due + " " + paylnk + " \\n\\n[This is a one-way message; we won't get replies.]",
            showCancelButton: true,
            confirmButtonText: 'Send',
            showLoaderOnConfirm: true,
            preConfirm: (message) => {
                postData = {
                    to: to,
                    message: message
                }
                let formBody = [];
                for (const property in postData) {
                    formBody.push(encodeURIComponent(property) + "=" + encodeURIComponent(postData[property]));
                }
                    
                return fetch(getPyScriptAddress() + "?a=send", {
                    method: 'POST',
                    body: formBody.join("&"),
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
                    }
                })
                .then(response => {
                    if (!response.ok) {
                        throw new Error(response.statusText)
                    }
                    return response.blob()
                })
                .catch(error => {
                    Swal.showValidationMessage(
                        `Request failed: ${error}`
                    )
                })
            },
            allowOutsideClick: () => !Swal.isLoading()
        }).then((result) => {
            if (result.isConfirmed) {
                Swal.fire({
                    title: 'Sent!',
                    showConfirmButton: false,
                    timer: 1500
                })
            }
        })
    }
            
            
            
    </script>
            
    <style>
    .attendee {
        margin: 0 0 1em;
        border: 1px solid #999;
        padding: 1em;
    }
        
    div.attendee h3 {
        margin-top: 0;
    }
    </style>   
''')
