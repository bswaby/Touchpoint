ProgramID = model.Data.ProgramID

for a in q.QuerySql("Select Name From Program Where Id = " + ProgramID):
    ProgramName = a.Name 
    
model.Header = ProgramName + ' | Program Manager'


familysql = '''
DECLARE @ProgramID int = ''' + ProgramID +'''
SELECT 
    dbo.People.PeopleId,
    dbo.People.Name, 
    dbo.People.FamilyId, 
    dbo.People.FirstName, 
    dbo.People.LastName, 
    dbo.People.Age,
    dbo.People.CellPhone,
    dbo.People.HomePhone,
    Organizations_alias1.OrganizationName,
    Organizations_alias1.OrganizationId,
	dbo.TransactionSummary.TotDue
FROM 
    dbo.ProgDiv 
    INNER JOIN dbo.Program ON (dbo.ProgDiv.ProgId = dbo.Program.Id) 
    INNER JOIN dbo.Division ON (dbo.ProgDiv.DivId = dbo.Division.Id) 
    INNER JOIN dbo.Organizations Organizations_alias1 ON (dbo.Division.Id = Organizations_alias1.DivisionId) 
    INNER JOIN dbo.OrganizationMembers ON (Organizations_alias1.OrganizationId = dbo.OrganizationMembers.OrganizationId) 
    LEFT JOIN dbo.TransactionSummary ON (dbo.OrganizationMembers.TranId = dbo.TransactionSummary.RegId) 
    INNER JOIN dbo.People ON (dbo.OrganizationMembers.PeopleId = dbo.People.PeopleId) 
    LEFT JOIN dbo.OrganizationExtra ON Organizations_alias1.OrganizationId = OrganizationExtra.OrganizationID
WHERE 
    dbo.Program.Id = @ProgramID --AND EXISTS (Select CheckInTimes_alias1.CheckInTime From dbo.CheckInTimes Where dbo.People.PeopleId = CheckInTimes_alias1.PeopleId)
    AND OrganizationExtra.Field = 'MemberManagerEnabled' 
    AND OrganizationExtra.BitValue = 1
    AND dbo.People.FamilyId = {0}
    AND dbo.OrganizationMembers.MemberTypeId = 220
GROUP BY Organizations_alias1.OrganizationId, 
    Organizations_alias1.OrganizationName, 
    dbo.People.Name, 
    dbo.People.FamilyId, 
    dbo.People.PeopleId, 
    dbo.People.FirstName, 
    dbo.People.LastName, 
    dbo.People.Age,
	dbo.TransactionSummary.TotDue,
    dbo.People.CellPhone,
    dbo.People.HomePhone
ORDER BY 
    dbo.People.LastName ASC, 
    dbo.People.FirstName ASC;
'''

listsql = '''
DECLARE @ProgramID int = ''' + ProgramID +'''
SELECT 
    Distinct dbo.People.FamilyId
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
    dbo.People.FamilyId
'''

subgrouplistsql = '''
DECLARE @ProgramID int = ''' + ProgramID +'''
SELECT 
    dbo.MemberTags.Name AS [SubGroup]
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
    AND dbo.People.PeopleId = {0}
GROUP BY 
    dbo.MemberTags.Name
    '''


print ('''
<script src ="https://code.jquery.com/jquery-3.5.1.min.js"></script>
<script src ="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"> </script> 
<!-- jQuery Modal -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css" />

<a href="'''+ model.CmsHost + '''/PyScript/MM-Charge?ProgramID=''' + ProgramID + '''"><i class="fa fa-usd fa-2x" aria-hidden="true"></i></a>
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
  border-collapse: collapse;
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
th, tr, td {  
  padding: 5px 5px 5px 10px;  
  padding-top: 1px;
  padding-bottom: 2px;
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

print '''<tr role = "row"><td style="background-color:#D3D3D3"></td>
    <td style="background-color:#D3D3D3"></td>
    <td style="background-color:#D3D3D3"></td>
    <td style="background-color:#D3D3D3"></td>'''

#get list of families
families = q.QuerySql(listsql)
for a in families:
    
    FamilyId = ""
    FamilyId = a.FamilyId
    
    #get each family member participating
    TuitionID = q.QuerySql(familysql.format(a.FamilyId))

    # print '''<tr role = "row"><td style="background-color:#D3D3D3"></td>
    #     <td style="background-color:#D3D3D3"></td>
    #     <td style="background-color:#D3D3D3"></td>
    #     <td style="background-color:#D3D3D3"></td>'''
    
    grandTotal = 0

    #interate through each family 
    for tID in TuitionID:
        print '<tr role = "row">'
        paylink = " "
        PayIDNote = " " 
        PayID = " "
        paylinkauth = " "
        hasusername = " "
        headCheck = " "
        userID = " " 
        userIDType = " " 
        AltPayID = 0




        #for each family identify a payer, set that variable, add up the total to be owed, have one payment link, figure out how to make one payment go across multiple people

        #on payment, maybe have some logic where balance applies to payer first, then trickles down to the subsequent family members

        #with any fees on child, move amounts from child to payer, therefore, all payments will not need to be split

        #at the time of payment, merge all outstanding balances for the family in that program to the payer. That way, each individual will have their specific total
        



        if model.ExtraValueInt(int(tID.PeopleId), str(ProgramID) + '_AltPayID') != 0:
            #Check to see if there is an alternate pay id
            AltPayID = model.ExtraValueInt(int(tID.PeopleId), str(ProgramID) + '_AltPayID')
            PayPerson = q.QuerySql("Select PeopleId, EmailAddress, FirstName, LastName, CellPhone, HomePhone from People Where PeopleId = "+ str(AltPayID))
            for alt in PayPerson:
                PayID = alt.PeopleId
                
                #set contact info for altPayID
                phoneContact = ""
                if alt.CellPhone != "":
                    phoneContact = model.FmtPhone(alt.CellPhone, " c: ")
                if alt.HomePhone != "":
                    phoneContact = phoneContact + model.FmtPhone(alt.HomePhone, " h:")
                    
                PayIDNote = PayIDNote + '''<i class="fa fa-usd" aria-hidden="true"></i>'''
                
                #check to see if altpayperson has a username
                hasusername = q.QuerySqlInt("SELECT COUNT(UserId) FROM Users Where PeopleId = " + str(alt.PeopleId)) 
                if hasusername == 0:
                    #add username so paylinks work
                    model.AddRole(alt.PeopleId,"Access")
                    model.RemoveRole(alt.PeopleId, "Access")
                
                #altPayID info
                PayIDNote = PayIDNote + '<i>alternate pay: ' + alt.FirstName + ' ' + alt.LastName + phoneContact + '</i><br>'


        #pull head of house
        Parents = q.QuerySql("Select PeopleId, EmailAddress, FirstName, LastName, CellPhone, HomePhone from dbo.People where FamilyId = " + str(FamilyId) + " AND PositionInFamilyId = 10")
        
        for parent in Parents:
            phoneContact = ""
            if parent.CellPhone != "":
                phoneContact = model.FmtPhone(parent.CellPhone, " c: ")
            if parent.HomePhone != "":
                phoneContact = phoneContact + model.FmtPhone(parent.HomePhone, " h:")
             
            if AltPayID == 0:
                #check to see if parent is head
                headCheck = q.QuerySqlInt("SELECT COUNT(FamilyId) FROM Families WHERE FamilyId = " + str(tID.FamilyId) + " AND HeadOfHouseholdId = " + str(parent.PeopleId))

                if headCheck != 0:
                    PayID = parent.PeopleId
                    PayIDNote = PayIDNote + '''<i class="fa fa-usd" aria-hidden="true"></i>'''
            
                
            #check to see username exists
            hasusername = q.QuerySqlInt("SELECT COUNT(UserId) FROM Users Where PeopleId = " + str(parent.PeopleId)) 
            if hasusername == 0:
                #add username so paylinks work
                model.AddRole(parent.PeopleId,"Access")
                model.RemoveRole(parent.PeopleId, "Access")

            PayIDNote = PayIDNote  + parent.FirstName + ' ' + parent.LastName + phoneContact + '<br>'
            
        if tID.TotDue != None:
            due = '${:,.2f}'.format(tID.TotDue)
            paylink = model.GetPayLink(PayID, tID.OrganizationId)
            try:
                paylinkauth = model.GetAuthenticatedUrl(tID.OrganizationId, paylink, True)
            except:
                paylinkauth = " "

        print '''<td role = "cell">''' 

        print '''
          <a href="{6}/PyScript/MM-MemberDetails?p1={1}&FamilyId={3}&ProgramName={4}&ProgramID={5}"> {0} ({2})</a>&nbsp
          <a href="{6}/Person2/{1}#tab-current" target="_blank"><i class="fa fa-info-circle" aria-hidden="true"></i></a>
          <br>{7}</td>'''.format(tID.Name, tID.PeopleId, tID.Age, tID.FamilyId, ProgramName,ProgramID,model.CmsHost, PayIDNote)

            
        print ('<td role = "cell"><a href="') + model.CmsHost + '/Org/{1}" target="_blank">{0}</a><br>'.format(tID.OrganizationName,tID.OrganizationId)

        #display subgroups the participant is in
        subGroupResults = q.QuerySql(subgrouplistsql.format(tID.PeopleId))
        for a in subGroupResults:
            print '{0}<br>'.format(a.SubGroup)

        print ('</td><td role = "cell">{0}</td>').format("%.2f" % float(tID.TotDue))


        #grab the grand total per family
        grandTotal = grandTotal + float(tID.TotDue)

        print '''<td>'''
        #Where CC and cash icon are
        if tID.TotDue != 0 or tID.TotDue != 0.0000:
            if paylinkauth != " ":
                print '<a href="MM-PaymentNotify?pid={0}&totaldue={1}&oid={2}&ProgramName={3}&ProgramID={4}&AltPayID={5}&FamilyId={6}"><i class="fa fa-credit-card-alt fa-3x" aria-hidden="true"></i></a>'.format(PayID, tID.TotDue, tID.OrganizationId,ProgramName,ProgramID,AltPayID,tID.FamilyId)
                print '''
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
                    <a href="#payfee{1}{2}" rel="modal:open"><i class="fa fa-money fa-4x" aria-hidden="true"></i></a></br>'''.format(tID.TotDue, tID.PeopleId, tID.OrganizationId,ProgramName,ProgramID)
        print '</td>'
        print '</tr>'
  
    print ('</td></tr>')

    print '''<tr><td>hello world</td><td></td><td>${0}</td><td>'''.format(grandTotal)

    if tID.TotDue != 0 or tID.TotDue != 0.0000:
        if paylinkauth != " ":
            print '<a href="MM-PaymentNotify?pid={0}&totaldue={1}&oid={2}&ProgramName={3}&ProgramID={4}&AltPayID={5}&FamilyId={6}"><i class="fa fa-credit-card-alt fa-3x" aria-hidden="true"></i></a>'.format(PayID, tID.TotDue, tID.OrganizationId,ProgramName,ProgramID,AltPayID,tID.FamilyId)
            print '''
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
                <a href="#payfee{1}{2}" rel="modal:open"><i class="fa fa-money fa-4x" aria-hidden="true"></i></a></br>'''.format(tID.TotDue, tID.PeopleId, tID.OrganizationId,ProgramName,ProgramID)
    print '</td>'
    print '</tr>'

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