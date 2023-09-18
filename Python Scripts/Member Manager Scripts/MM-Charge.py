ProgramID = model.Data.ProgramID

#ProgramName = model.Data.ProgramName

for a in q.QuerySql("Select Name From Program Where Id = " + ProgramID):
    ProgramName = a.Name #model.Data.ProgramName
    
model.Header = ProgramName + ' | Program Charge Manager'

chargeVariables = {
    'fees' : [
    #Peer Place Defined Fees
    {
        'id': "1143_2-Day",
        'name': "2-Day",
        'frequency': "Monthly",
        'description': "2-Day a Week Fee",
        'cost': 200
    },
    {
        'id': "1143_3-Day",
        'name': "3-Day",
        'frequency': "Monthly",
        'description': "3-Day a Week Fee",
        'cost': 300
    },
    {
        'id': "1143_4-Day",
        'name': "4-Day",
        'frequency': "Monthly",
        'description': "4-Day a Week Fee",
        'cost': 400
    },
        #Rooted Defined Fees
    {
        'id': "1149_2-Day",
        'name': "2-Day",
        'frequency': "Monthly",
        'description': "Monthly Tuition Fee",
        'cost': 80
    },
    #Test Defined Fees
    {
        'id': "1152_2-Day",
        'name': "2-Day",
        'frequency': "Monthly",
        'description': "2-Day a Week Fee",
        'cost': 200
    },
    {
        'id': "1152_3-Day",
        'name': "3-Day",
        'frequency': "Monthly",
        'description': "3-Day a Week Fee",
        'cost': 300
    },
    {
        'id': "1152_4-Day",
        'name': "4-Day",
        'frequency': "Monthly",
        'description': "4-Day a Week Fee",
        'cost': 400
    }],
    'discounts' : 
    [{
            'title': "Staff Discount",
            'percentage': .50,
            'amount': 50,
            'discountType': "percentage",
            'code': 1,
        },{
            'title': "Multiple Family Member Discount",
            'percentage': .10,
            'amount': 2,
            'discountType': "amount",
            'code': 2,
        }
    ]
}


familysql = '''
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
    dbo.People.EmailAddress, 
    dbo.People.Age,
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



#page styling
print '''
<head>

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
<script src ="https://code.jquery.com/jquery-3.5.1.min.js"></script>                 
<script src ="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"> </script>  
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.0.0/jquery.min.js"></script>

<!-- jQuery Modal -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css" />

	<style>
		button, #buttonLink {
            width: 45%;
			color: #ffffff;
			background-color: #2d63c8;
			font-size: 15px;
			border: 1px solid #2d63c8;
			padding: 5px 5px;
			cursor: pointer;
			display: inline-block;
			float: left;
            text-decoration: none;
            text-align: center;
            margin-right: 5px;
		}
		button:hover, #buttonLink:hover {
			color: #2d63c8;
			background-color: #ffffff;
		}
	</style>
</head>
'''

#<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">


print '''
<table width="100%">
    <tr>
        <td align="left"><a href="''' + model.CmsHost + '''/PyScript/MM-MemberManager?ProgramID=''' + ProgramID + '''"><i class="fa fa-home fa-3x"></i></a></td>
        <td align="right"><a href="''' + model.CmsHost + '''/PyScript/MM-ChargeAll?ProgramID=''' + ProgramID + '''">Charge All</a></td>
    </tr>
</table>
    <br>
  <div class = "table-responsive">  
    <table role = "table" class = "table filtered-table">  
      <thead role = "rowgroup">  
        <tr role = "row">  
          <th role = "columnheader"> Name </th>  
          <th role = "columnheader"> Involvement (subgroup) </th>
          <th role = "columnheader"> Outstanding </th>
          <th role = "columnheader"> New Charges </th>
        </tr>  
      </thead>  
      <tbody role = "rowgroup">  

'''

# for b in q.QuerySql(sqlOrganizations):
#     organizationID = b.OrganizationId

organizationID = 0
for a in q.QuerySql(listsql):

    #get each family member participating
    TuitionID = q.QuerySql(familysql.format(a.FamilyId))

    #simple way to group like families together visually
    print '''<tr role = "row"><td style="background-color:#D3D3D3"></td>
        <td style="background-color:#D3D3D3"></td>
        <td style="background-color:#D3D3D3"></td>
        <td style="background-color:#D3D3D3"></td>'''
    
    #pull tuition
    for tID in TuitionID:
        organizationID = tID.OrganizationId
        print '<tr role = "row">'
        print '''<td role = "cell">
          <a href="{12}/PyScript/MM-MemberDetails?p1={1}&FamilyId={3}&ProgramName={4}&ProgramID={5}"> {0} ({2})</a>
             &nbsp<a href="{12}/Person2/{1}#tab-current" target="_blank"><i class="fa fa-info-circle" aria-hidden="true"></i></a>
          </td>'''.format(tID.Name, tID.PeopleId, tID.Age, tID.FamilyId, ProgramName, ProgramID,tID.EmailAddress,tID.PrimaryAddress,tID.PrimaryCity,tID.PrimaryState,tID.PrimaryZip,tID.CellPhone,model.CmsHost)
        
        #pull all subgroup(s) assigned
        subGroupList = q.QuerySql(subgrouplistsql.format(tID.PeopleId))
        # changes needed here to implement subgrouplistsql
        print ('<td role = "cell"><a href="') + model.CmsHost + '/Org/{0}" target="_blank">'.format(tID.OrganizationId) + '{0}</a><br>'.format(tID.OrganizationName)

        subGroupResults = q.QuerySql(subgrouplistsql.format(tID.PeopleId))
        i = 0
        for a in subGroupResults:
            if len(subGroupResults)>1 & i != (len(subGroupResults)-1):
                i = i+1
            else:
                print(a.SubGroup)
            
        if tID.TotDue == None:
            tID.TotDue = 0.00

        print '</td><td role = "cell">{0}</td>'.format("%.2f" % float(tID.TotDue))

        cost = 0
        totalDiscountPercentage = 0
        totalDiscountAmount = 0

        for sgl in subGroupList:
        
            for cs in chargeVariables['fees']:
                if (str(ProgramID) + "_"+ str(sgl.SubGroup)) == cs['id']:
                    cost = cs['cost']
            
            discountCode = 1
            for vc in chargeVariables['discounts']:
                if discountCode == vc['code']:
                    if sgl.SubGroup == "Staff":
                        if vc['discountType'] == "percentage":
                            totalDiscountPercentage = totalDiscountPercentage + vc['percentage']
                        else:
                            totalDiscountAmount = totalDiscountAmount + vc['amount']

        discountCode = 2
        for vc in chargeVariables['discounts']:
            if discountCode == vc['code']:
                if len(TuitionID) > 1:
                    if vc['discountType'] == "percentage":
                            totalDiscountPercentage = totalDiscountPercentage + vc['percentage']
                    else:
                        totalDiscountAmount = totalDiscountAmount + vc['amount']
        
        cost = cost - (cost*totalDiscountPercentage)
        cost = cost - totalDiscountAmount

        if cost < 0:
            cost = 0

        formatCost = "%.2f" % cost

        # model.AddTransaction(int(tID.PeopleId), int(paymentOrg.OrganizationId), float("%.2f" % cost), messageDescription)

        print '''<td role = "cell">
        <form id="chargeIndividual" action="MM-ChargeIndividual">
                  <div>
                   <input type="hidden" id="pid" name="pid" value="{1}">
                   <input type="hidden" id="PaymentOrg" name="PaymentOrg" value="{2}">
                    <input type="hidden" id="ProgramID" name="ProgramID" value="{3}">
                    <input type="hidden" name="PayAmount" value="{0}"/>
                    <input type="hidden" name="PaymentType" value="FEE">
                    <input type="hidden" name="PaymentDescription" id="PaymentDescription" value="This is a charge of ${0} for {4}"/>
                   </div>
                  <button>Charge ${5} Now</button>
                </form>
     
        
        <a id = "buttonLink" href="#chargeIndividualVariable{1}" rel="modal:open">Variable Charge</a>
            <form id="chargeIndividualVariable{1}" class = "modal" action="MM-ChargeIndividual">
                <div>
                    <input type="hidden" id="pid" name="pid" value="{1}">
                    <p>{1}</p>
                    <input type="hidden" id="PaymentOrg" name="PaymentOrg" value="{2}">
                    <input type="hidden" id="ProgramID" name="ProgramID" value="{3}">
                    <input type="hidden" name="PaymentType" value="FEE" id="PaymentType"/>
                    <input type="hidden" name="PaymentDescription" id="PaymentDescription" value="This is a charge for {4}"/>
                    <div class="modalparagraph">
                    <input type="number" name="PayAmount" step="any" id="PayAmount"/> &nbsp
                    
                    </div>
                  
                  </div>
                  <button>Variable Charge</button> 
                </form>
            </form>
            </td>'''.format(formatCost, tID.PeopleId, organizationID, ProgramID, ProgramName, formatCost)

#TODO add modal in for variable charge
#paymentOrg.OrganizationId

        print '</td>'
        print '</tr>'
    
    print ('</td></tr>')