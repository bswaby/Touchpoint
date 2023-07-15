ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

FamilyId = model.Data.FamilyId

model.Header = ProgramName +' Member Details'

declareVariables = '''
DECLARE @ProgramID int = ''' + ProgramID + '''
DECLARE @ProgramName varchar(30) = ''' + ProgramName

q.QuerySql(declareVariables)

listsql = '''
    SELECT 
        dbo.People.HomePhone, 
        dbo.People.CellPhone, 
        dbo.People.Name, 
        dbo.People.FamilyId, 
        dbo.People.PeopleId, 
        dbo.People.PositionInFamilyId, 
        dbo.People.FirstName, 
        dbo.People.LastName, 
        dbo.People.EmailAddress, 
        dbo.People.PrimaryAddress, 
        dbo.People.PrimaryCity, 
        dbo.People.PrimaryState, 
        dbo.People.PrimaryZip, 
        dbo.People.Age
    FROM 
        dbo.People
    WHERE 
        dbo.People.PeopleId = @P1
'''


transactions = '''
    SELECT 
        Organizations_alias1.OrganizationName, 
        Organizations_alias1.OrganizationId, 
        dbo.People.HomePhone, 
        dbo.People.CellPhone, 
        dbo.People.Name, 
        dbo.TransactionSummary.TotDue AS [TotalDue],
        dbo.TransactionSummary.TranDate,
        FORMAT(dbo.TransactionSummary.TotPaid, 'C') AS [TotalPaid],
        dbo.People.FamilyId, 
        dbo.People.PeopleId, 
        dbo.People.PositionInFamilyId, 
        dbo.People.FirstName, 
        dbo.People.LastName, 
        dbo.People.EmailAddress, 
        dbo.People.PrimaryAddress, 
        dbo.People.PrimaryCity, 
        dbo.People.PrimaryState, 
        dbo.People.PrimaryZip, 
        dbo.People.Age, 
        dbo.TransactionSummary.RegId
    FROM 
        dbo.OrganizationMembers 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.OrganizationMembers.PeopleId = dbo.People.PeopleId) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.OrganizationMembers.OrganizationId = Organizations_alias1.OrganizationId) 
    INNER JOIN 
        dbo.TransactionSummary 
    ON 
        ( 
            dbo.OrganizationMembers.TranId = dbo.TransactionSummary.RegId) 
    WHERE 
        dbo.People.PeopleId = @P1--3134;
    ORDER BY 
        dbo.TransactionSummary.TranDate Desc
    '''

sqlemails = '''
    SELECT Top 20
            dbo.People.Name, 
            dbo.EmailQueue.Subject, 
            dbo.EmailQueue.Body, 
            dbo.People.PeopleId, 
            dbo.EmailQueue.Sent, 
            dbo.EmailQueue.FromName 
        FROM 
            dbo.EmailQueueTo
        INNER JOIN 
            dbo.EmailQueue 
        ON 
            ( 
                dbo.EmailQueueTo.Id = dbo.EmailQueue.Id) 
        INNER JOIN 
            dbo.People 
        ON 
            ( 
                dbo.EmailQueueTo.PeopleId = dbo.People.PeopleId)
        Where dbo.EmailQueueTo.PeopleId = @P1 --39337
        Order by dbo.EmailQueue.Sent Desc
'''

sqlemails1 = '''
    SELECT Top 20
        dbo.People.Name, 
        dbo.EmailQueue.Subject, 
        dbo.EmailQueue.Body, 
        dbo.People.PeopleId, 
        dbo.EmailQueue.Sent AS [DateSent], 
        dbo.EmailQueue.FromName 
    FROM 
        dbo.EmailQueueTo 
    INNER JOIN 
        dbo.EmailQueue 
    ON 
        ( 
            dbo.EmailQueueTo.Id = dbo.EmailQueue.Id) 
    INNER JOIN 
        dbo.People 
    ON 
        ( 
            dbo.EmailQueue.QueuedBy = dbo.People.PeopleId)
    Where dbo.People.PeopleId =  @P1 --39337
    Order by dbo.EmailQueue.Sent Desc
    '''

sqlMemberOrganization = '''SELECT 
        Organizations_alias1.OrganizationName,
        Organizations_alias1.OrganizationId,
        dbo.OrganizationMembers.MemberTypeId,
        dbo.OrganizationMembers.PeopleId
    FROM 
        dbo.OrganizationMembers 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.OrganizationMembers.OrganizationId = Organizations_alias1.OrganizationId) 
    WHERE 
        dbo.OrganizationMembers.PeopleId = @P1
        --AND dbo.OrganizationMembers.MemberTypeId <> 220 --3134
    Order By Organizations_alias1.OrganizationName;
'''

sqlOrganizations = '''
    SELECT 
        Organizations_alias1.OrganizationId,
        Organizations_alias1.OrganizationName
    FROM 
        dbo.ProgDiv 
    INNER JOIN 
        dbo.Program 
    ON 
        ( 
            dbo.ProgDiv.ProgId = dbo.Program.Id) 
    INNER JOIN 
        dbo.Division 
    ON 
        ( 
            dbo.ProgDiv.DivId = dbo.Division.Id) 
    INNER JOIN 
        dbo.Organizations Organizations_alias1 
    ON 
        ( 
            dbo.Division.Id = Organizations_alias1.DivisionId) 
    INNER JOIN 
        dbo.OrganizationMembers 
    ON 
        ( 
            Organizations_alias1.OrganizationId = dbo.OrganizationMembers.OrganizationId) 
    WHERE 
        dbo.Program.Id = ''' + ProgramID + ''' --and dbo.OrganizationMembers.PeopleId = 3134
    GROUP BY Organizations_alias1.OrganizationId, 
        Organizations_alias1.OrganizationName
    ORDER BY 
        Organizations_alias1.OrganizationName
    '''

sqlCheckinTimes1 = '''
    SELECT 
        CheckInTimes_alias1.CheckInTime 
    FROM 
        dbo.CheckInTimes CheckInTimes_alias1 
    WHERE 
        CheckInTimes_alias1.PeopleId = @P1
    AND CheckInTimes_alias1.location = @ProgramName 
    ORDER BY 
        CheckInTimes_alias1.CheckInTime DESC;
'''

sqlCheckinTimes = '''
    SELECT 
        CheckInTimes_alias1.CheckInTime,
        CheckInTimes_alias1.location
    FROM 
        dbo.CheckInTimes CheckInTimes_alias1 
    WHERE 
        CheckInTimes_alias1.PeopleId = @P1
    --AND CheckInTimes_alias1.location = @ProgramName 
    ORDER BY 
        CheckInTimes_alias1.CheckInTime DESC;
'''

sqlPayments = '''
    Select 
        t.Id,
        t.TransactionDate,
        t.TransactionId,
        REPLACE(
            REPLACE(
            REPLACE(t.message, 'APPROVED', 'Online Transaction')
                   , 'Response: ',  '')
                   ,'CC -', 'CC |') as Description,
        FORMAT(t.amt, 'C') AS [amt],
        t.AdjustFee,
        ts.PeopleId,
        ts.OrganizationId,
        org.OrganizationName,
        ts.RegId
    From TransactionSummary ts
    Left Join [Transaction] t
    ON t.originalId = ts.regid 
    Left Join Organizations org
    On ts.OrganizationId = org.OrganizationId
    where 
      PeopleId = @P1
      AND amt <> 0
      AND (AdjustFee is NULL  OR AdjustFee = 0)
    Order by TransactionDate'''
    
sqlPaymentsNew = '''
SELECT 
    t.Id, 
    t.TransactionDate, 
    t.TransactionId,         
    REPLACE(
            REPLACE(
            REPLACE(t.message, 'APPROVED', 'Online Transaction')
                   , 'Response: ',  '')
                   ,'CC -', 'CC |') as Description,
    FORMAT(t.amt, 'C') AS [amt],
    t.AdjustFee,
    tp.PeopleId,
    tp.OrgId,
    org.OrganizationName
    --tp.RegId 
FROM [Transaction] t
LEFT JOIN [TransactionPeople] tp ON t.OriginalId = tp.Id
LEFT JOIN Organizations org On t.OrgId = org.OrganizationId
WHERE (tp.PeopleId = @P1) AND (t.AdjustFee = 0 OR t.TransactionGateway <> ' ') AND t.amt <> 0
'''

sqlFamily = '''
    SELECT 
        dbo.People.Name, 
        dbo.People.FamilyId, 
        dbo.People.PeopleId, 
        lookup.FamilyPosition.Description AS [FamilyPosition], 
        dbo.People.FirstName, 
        dbo.People.LastName, 
        dbo.People.EmailAddress, 
        dbo.People.Age,
        dbo.Picture.SmallUrl 
    FROM 
        dbo.People 
    LEFT JOIN 
        lookup.FamilyPosition 
    ON 
        ( 
            dbo.People.PositionInFamilyId = lookup.FamilyPosition.Id) 
    LEFT JOIN 
        dbo.Picture 
    ON 
        ( 
            dbo.People.PictureId = dbo.Picture.PictureId) 
    WHERE 
        dbo.People.FamilyId = {};
'''.format(FamilyId)

print '''
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
<script src ="https://code.jquery.com/jquery-3.5.1.min.js"></script>                 
<script src ="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"> </script>  
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.0.0/jquery.min.js"></script>

<!-- jQuery Modal -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css" />


<a onclick="history.back()"><i class="fa fa-hand-o-left fa-2x" aria-hidden="true"></i></a>&nbsp;&nbsp;
<a href="'''+ model.CmsHost + '''/PyScript/MM-MemberManager"><i class="fa fa-home fa-2x"></i></button>


<style>    
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
  background: #fff;  
}  
th {  
  color: #ffffff;  
  background: #39a2fb;  
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
  background: #fff; 
}  
td {  
  padding: 5px 5px 5px 10px;  
  word-wrap: break-word;  
  border-bottom: 1px solid #ccc;  
  background: #fff;  
}  
 
.headshot {
    border-radius: 999999px;
    height: 35px;
    width: 35px;
    display: inline-block;
    background-repeat: no-repeat;
    background-size: cover;
    float: left;
    margin-right: 10px;
}

.profile-photo {
    background: #fff;  
}
#profile-portrait {
    height: 180px;
    width: 180px;
}
a.bottom-right-picture {
    position: absolute;
    left: 170px;
    top: 170px;
}

.table-responsive {
    width: 80%;
    margin-bottom: 15px;
    overflow-y: hidden;
    -ms-overflow-style: -ms-autohiding-scrollbar;
    border: 1px solid #ddd;
}
.table-responsive {
    min-height: .01%;
    overflow-x: auto;
}
.table {
    width: 100%;
    max-width: 100%;
    margin-bottom: 20px;
    background: #fff; 
}
table {
    background-color: transparent;
}
table {
    border-collapse: collapse;
    border-spacing: 0;
    background: #fff; 
}
.fade {
    opacity: 0;
    -webkit-transition: opacity .15s linear;
    -o-transition: opacity .15s linear;
    transition: opacity .15s linear;
}
.nav-tabs, .nav-pills {
    position: relative;
}

[class*=col-] {
    padding-top: 0px;
    padding-bottom: 0px;
}

@media (max-width: 767px)
.col-lg-1, .col-lg-10, .col-lg-11, .col-lg-12, .col-lg-2, .col-lg-3, .col-lg-4, .col-lg-5, .col-lg-6, .col-lg-7, .col-lg-8, .col-lg-9, .col-md-1, .col-md-10, .col-md-11, .col-md-12, .col-md-2, .col-md-3, .col-md-4, .col-md-5, .col-md-6, .col-md-7, .col-md-8, .col-md-9, .col-sm-1, .col-sm-10, .col-sm-11, .col-sm-12, .col-sm-2, .col-sm-3, .col-sm-4, .col-sm-5, .col-sm-6, .col-sm-7, .col-sm-8, .col-sm-9, .col-xs-1, .col-xs-10, .col-xs-11, .col-xs-12, .col-xs-2, .col-xs-3, .col-xs-4, .col-xs-5, .col-xs-6, .col-xs-7, .col-xs-8, .col-xs-9 {
    margin-bottom: 10px;
}
.col-lg-1, .col-lg-10, .col-lg-11, .col-lg-12, .col-lg-2, .col-lg-3, .col-lg-4, .col-lg-5, .col-lg-6, .col-lg-7, .col-lg-8, .col-lg-9, .col-md-1, .col-md-10, .col-md-11, .col-md-12, .col-md-2, .col-md-3, .col-md-4, .col-md-5, .col-md-6, .col-md-7, .col-md-8, .col-md-9, .col-sm-1, .col-sm-10, .col-sm-11, .col-sm-12, .col-sm-2, .col-sm-3, .col-sm-4, .col-sm-5, .col-sm-6, .col-sm-7, .col-sm-8, .col-sm-9, .col-xs-1, .col-xs-10, .col-xs-11, .col-xs-12, .col-xs-2, .col-xs-3, .col-xs-4, .col-xs-5, .col-xs-6, .col-xs-7, .col-xs-8, .col-xs-9 {
    position: relative;
    min-height: 1px;
    padding-right: 15px;
    padding-left: 15px;
}
@media (max-width: 768px)
.tab-content {
    border: 0!important;
    padding-left: 0!important;
    padding-right: 0!important;
}
.tab-content {
    background-color: #fff;
    border-left: 1px solid #ddd;
    border-right: 1px solid #ddd;
    border-bottom: 1px solid #ddd;
    padding-top: 5px;
    padding-left: 5px;
    padding-right: 5px;
}
*, :after, :before {
    -webkit-box-sizing: border-box;
    -moz-box-sizing: border-box;
    box-sizing: border-box;
}
.modal {
    display: none;
    vertical-align: middle;
    position: relative;
    z-index: 2;
    max-width: 600px;
    box-sizing: border-box;
    width: 90%;
    background: #fff;
    padding: 15px 30px;
    -webkit-border-radius: 8px;
    -moz-border-radius: 8px;
    -o-border-radius: 8px;
    -ms-border-radius: 8px;
    border-radius: 8px;
    -webkit-box-shadow: 0 0 10px #000;
    -moz-box-shadow: 0 0 10px #000;
    -o-box-shadow: 0 0 10px #000;
    -ms-box-shadow: 0 0 10px #000;
    box-shadow: 0 0 10px #000;
    text-align: left;
}
div {
}
#modalparagraph {
    visibility: visible;
}
.attendee {
    margin: 0 0 1em;
    border: 1px solid #999;
    padding: 1em;
}
    
div.attendee h3 {
    margin-top: 0;
}
</style>  
<body>  

        
<div class="container-fluid" id="main">
        
    <br />
  
        <div class="col-sm-4 col-md-3 col-lg-2 hidden-print">

        <div id="sidebar">
            <div class="row">
                <div class="col-sm-12">            
                    <a id="family-members-collapse" data-toggle="collapse" href="#family-members-section" aria-expanded="true" aria-controls="family-members-section"><i class="fa fa-chevron-circle-down"></i></a> <a href="">
                        Family Members
                    </a>
                </div>
            </div>  
            <div id="family-div">
                <div class="row collapse in" id="family-members-section">
                    <div class="col-sm-12">
                        <ul id="family_members" class="nav nav-stacked nav-tabs nav-tabs-left" style="margin-bottom: 10px;">
'''
for a in q.QuerySql(sqlFamily):
    print '<li class="active " style="font-size: 0.85em;">'
    print '<a href="https://myfbch.com/PyScript/MM-MemberDetails?p1={0}&FamilyId={1}&ProgramName={2}&ProgramID={3}" target="_blank">'.format(a.PeopleId,a.FamilyId, ProgramName, ProgramID)
    print '<div class="headshot" style="background-image:url({}); background-position: top"></div>'.format(a.SmallUrl)
    print '<span class="name">{0} {1} </span><br />'.format(a.FirstName,a.LastName)
    print '<span class="meta">'
    print '<span class="age">{}</span>'.format(a.Age)
    print '&bull; <span class="status"></span>'
    print '&bull; <span class="role">{}</span>'.format(a.FamilyPosition)
    print '''</span><div class="email email_display truncate"></div></a></li>'''
                        
print '''    </ul>
        </div>
    </div>
</div>
                    

        </div>
        </div>
    
        <div class="col-sm-8 col-md-9 col-lg-10">

</br>
<table width="100%">
<tr>
'''

for a in q.QuerySql("SELECT Top 1 dbo.Picture.SmallUrl FROM dbo.People INNER JOIN dbo.Picture ON (dbo.People.PictureId = dbo.Picture.PictureId) WHERE dbo.People.PeopleId = @P1"):
    print '''<td><div class="profile-photo"><img src="''' + a.SmallUrl + '" alt="Invisible Person or Missing Pic">'

for a in q.QuerySql(listsql):
    pid = a.PeopleId
    print '<h1 id="nameline" class="" style="margin-top:0"> {} </h1>'.format(a.Name)
    print '<a href="' + model.CmsHost + '/Person2/{0}#tab-current" target="_blank"><i class="fa fa-info-circle" aria-hidden="true"></i></a>'.format(a.PeopleId)
    print '''<ul class="meta list-unstyled">
        <li>
            <div class="profile-photo">
                            <i class="fa fa-home fa-fw"></i>'''
    print '<a href="#" class="dropdown-toggle " data-toggle="dropdown"> {}, {}, {} {}</a>&nbsp;&nbsp;'.format(a.PrimaryAddress, a.PrimaryCity, a.PrimaryState, a.PrimaryZip)
    print '''</div>
                <div class="dropdown">
                </div>
                    <i class="fa fa-envelope fa-fw"></i>'''
    print '<span id="contactline"><a href="mailto:{0}" target="_blank">{0}</a> • {1} {2}</span>'.format(a.EmailAddress, model.FmtPhone(a.CellPhone, "c:"), model.FmtPhone(a.HomePhone, "h:"))
    print '</li></ul></div></td>'


print '''
    </tr>
    </table>
    <ul class="nav nav-tabs" id="person-tabs">
        <li class="active">
            <a href="#feesummary" aria-controls="feesummary" data-toggle="tab" aria-expanded="true">Fee Summary</a>
        </li>
        <li>
            <a href="#paymenthistory" aria-controls="paymenthistory" data-toggle="tab" aria-expanded="true">Payment History</a>
        </li>
        <li>
            <a href="#organization" aria-controls="organization" data-toggle="tab">Involvements</a>
        </li>
        <li>
            <a href="#notifications" aria-controls="notificatons" data-toggle="tab">Notifications (Last 20)</a>
        </li>
        <li id="involvementstop">
            <a href="#attendence" aria-controls="attendence" data-toggle="tab">Gym Attendance</a>
        </li>
    </ul>
    
    <div class="tab-content">
    <div class="tab-pane fade in active" id="feesummary">
     <section>   
      <div>  
        <table class="table-striped">  
          <thead role = "rowgroup">  
            <tr role = "row">  
              <th role = "columnheader"> Fee Group </th>  
              <th role = "columnheader"> Due </th>
              <th role = "columnheader"> Action </th> 
            </tr>  
          </thead>  
          <tbody role = "rowgroup">  '''

for a in q.QuerySql(transactions):
    paylink = " "
    paylinkauth = " "
    sendemail = "n"
    sendtext = "n"
    if a.TotalDue != None:
        due = '${:,.2f}'.format(a.TotalDue)
        #if a.OrganizationId != None and a.PeopleId != None:
        paylink = model.GetPayLink(a.PeopleId, a.OrganizationId)
        paylinkauth = model.GetAuthenticatedUrl(a.OrganizationId, paylink, True)
        #paylink = a.PeopleId + " : " + a.OrganizationID

    else:
        due = 0
        
    print '<tr role = "row">'
    print '<td role = "cell"><a href="' + model.CmsHost + '/Org/{1}" target="_blank"> {0} </a>( {2} )</td>'.format(a.OrganizationName, a.OrganizationId, a.TranDate) 
    print '<td role = "cell"> {} </td>'.format(due) 
    print '<td role = "cell">'
    if a.EmailAddress != None:
        sendemail = "y"
    
    if a.CellPhone != None:
        sendtext = "y"
    if due != 0:
        if paylinkauth != " ":
            print '<a href="PaymentNotify?pid={0}&totaldue={1}&oid={2}&sendemail={3}&sendtext={4}">Text/Email Payment Link</a></br>'.format(a.PeopleId, a.TotalDue, a.OrganizationId,sendemail,sendtext)
            print '''
                <form id="payfee{3}" class="modal" action="Payment">
                  <div class="modalparagraph">
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
                  <button >submit</button>
                </form>
                <a href="#payfee{3}" rel="modal:open">Pay by Cash/Check</a></br>'''.format('${:,.2f}'.format(a.TotalDue), a.PeopleId, a.OrganizationId, a.RegId)
            print '<a href="{0}" target="_blank">Pay in Person</a><i>(Open via Incognito)</i>'.format(paylink)
        
    #if a.TotalPaid != None:
    #    print '''</br><a href="''' + model.CmsHost + '''/PyScript/MM-Receipt?p1={}">Receipt</a>'''.format(a.RegId)    
    
    print '</td></tr>'

print ''' </tbody>  
            </table>  
        </div>
        </section>
        </div>
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
            <div class="tab-pane fade" id="paymenthistory">
        <section>   
          <div>  
            <table class="table-striped">  
              <thead role = "rowgroup">  
                <tr role = "row">  
                  <th role = "columnheader"> Transaction Date </th>
                  <th role = "columnheader"> Description </th>
                  <th role = "columnheader"> Involvement </th>
                  <th role = "columnheader"> Amount </th>
                  <th role = "columnheader"> Receipt </th>
                </tr>  
              </thead>  
            <tbody role = "rowgroup">'''

for a in q.QuerySql(sqlPaymentsNew):
    #count = count + 1
    print '<tr role = "row">'
    print '<td role = "cell"> {0} </td>'.format(a.TransactionDate)
    print '<td role = "cell"> {0} </td>'.format(a.Description)
    print '<td role = "cell"> {0} </td>'.format(a.OrganizationName)
    print '<td role = "cell"> {0} </td>'.format(a.amt)
    print '''<td role = "cell"><a href="''' + model.CmsHost + '''/PyScript/MM-Receipt?p1={0}&TranId={1}">Receipt</a></td>'''.format(a.PeopleId,a.Id)  
    print '</tr>'

print '''
          </tbody>  
         </table>
        </div>
        </section>  
        </div>   
        '''
        
print '<div class="tab-pane fade" id="organization">'

#add drop/add org features
print '''
    <form id="frmaddgroup" class="modal" action="AddToGroup">
          <p class="modalparagraph">
           <h3>Add to ''' + ProgramName + ''' Group</h3></br>
            <select name="addorg" id="addorg">
'''

for a in q.QuerySql(sqlOrganizations):
    print '<option value="{1}">{0}</option>'.format(a.OrganizationName, a.OrganizationId)
    
    
print '<input type="hidden" id="{0}" name="pid" value="{0}">'.format(pid)

print '''
        </select>
        </br>
        </br>
        <button >submit</button>
        </form>
        <a href="#frmaddgroup" rel="modal:open">Add to ''' + ProgramName + ''' Group</a></br>

         <section>  
        <div>  
            <table class="table-striped">  
              <thead role = "rowgroup">  
                <tr role = "row">  
                  <th role = "columnheader"> Involvements Currently In: </th>
                  <th role = "columnheader">  </th>  
                </tr>  
              </thead>  
              <tbody role = "rowgroup">  '''
addfee = 0
for a in q.QuerySql(sqlMemberOrganization):
    addfee = addfee + 1
    print '<tr role = "row">'
    print '<td role = "cell"> {} </td>'.format(a.OrganizationName)

    print '''
        <td role = "cell">
        <form id="payfee{2}" class="modal" action="Payment">
          <div class="modalparagraph">
           <input type="hidden" id="pid" name="pid" value="{0}">
           <input type="hidden" id="PaymentOrg" name="PaymentOrg" value="{1}">
           <input type="hidden" id="addpayment" name="addpayment" value="y">
          FEE Type:
           <input type="radio" name="PaymentType" value="FEE|" id="PaymentType">ADD FEE
           <br>
           <i>NOTE: FEE's <u><b>MUST</b></u> BE A NEGATIVE NUMBER</i>
          </div>
          <div class="modalparagraph">
          Description:
           <input type="text" name="PaymentDescription" id="PaymentDescription"/>
          </div>
          <div class="modalparagraph">
          Fee Amount:<input type="number" name="PayAmount" step="any" id="PayAmount"/>
          </div>
          <button >submit</button>
        </form>
        <a href="#payfee{2}" rel="modal:open">Add Fee</a></td>'''.format(a.PeopleId, a.OrganizationId, addfee)

    print '</tr>'

print '''
          </tbody>  
            </table>  
        </div> 
        </section>

        </div>
        <div class="tab-pane fade" id="notifications">
            <section>   
       <div>  
         <table class="table-striped">   
                  <thead role = "rowgroup">  
                    <tr role = "row">  
                      <th role = "columnheader"> Sent </th>
                      <th role = "columnheader"> From </th>
                      <th role = "columnheader"> Subject </th>
                      <th role = "columnheader"> Message </th>
                    </tr>  
                  </thead>  
                <tbody role = "rowgroup">  '''

count = 0            
for a in q.QuerySql(sqlemails):
    count = count + 1
    print '<tr role = "row">'
    print '<td role = "cell"> {0} </td>'.format(a.Sent)
    print '<td role = "cell"> {0} </td>'.format(a.FromName) 
    print '<td role = "cell"> {0} </td>'.format(a.Subject)
    print '''<td role = "cell">
      <form id="emailBody{1}" class="modal">
        <div class="modalparagraph">
         {0}
        </div>
      </form>
      <a href="#emailBody{1}" rel="modal:open"><i class="fa fa-file-image-o" aria-hidden="true"></i></a></td>'''.format(a.Body,count) 
    print '</tr>'
    
print '''
      </tbody>  
        </table>
    </div>
    </section>  
    </div> 

    <div class="tab-pane fade" id="attendence">
        <section>   
          <div>  
            <table class="table-striped">  
              <thead role = "rowgroup">  
                <tr role = "row">  
                  <th role = "columnheader"> Attended </th>
                </tr>  
              </thead>  
            <tbody role = "rowgroup">'''
            
            
for a in q.QuerySql(sqlCheckinTimes):
    print '<tr role = "row">'
    print '<td role = "cell"> {} </td>'.format(a.CheckInTime)
    print '<td role = "cell"> {} </td>'.format(a.location)
    print '</tr>'
    
print '''
          </tbody>  
         </table>
        </div>
        </section>  
        </div> 
        </div>
        </div>
        </div>
      
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
           

            
            
            
            
'''