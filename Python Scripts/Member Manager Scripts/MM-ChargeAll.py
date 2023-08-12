ProgramID = model.Data.ProgramID

#ProgramName = model.Data.ProgramName

for a in q.QuerySql("Select Name From Program Where Id = " + ProgramID):
    ProgramName = a.Name #model.Data.ProgramName
    
model.Header = ProgramName + ' | Program Charge All'

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


for b in q.QuerySql(sqlOrganizations):
    organizationID = b.OrganizationId


for a in q.QuerySql(listsql):

    #get each family member participating
    TuitionID = q.QuerySql(familysql.format(a.FamilyId))
    
    #pull tuition
    for tID in TuitionID:
        #pull all subgroup(s) assigned
        subGroupList = q.QuerySql(subgrouplistsql.format(tID.PeopleId))

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
        formatCost = "%.2f" % cost




        #make payment
        if tID.PeopleId == "":
            print ""
        elif cost == "":
            print '<h2>missing payment</h2>'
        elif organizationID == "":
            print '<h2>organization missing</h2>'
        else:
            messageDescription = "The all participants in " + ProgramName + " have been charged."
            model.AdjustFee(int(tID.PeopleId), int(organizationID), -float(cost), messageDescription)
            #model.Email(3134,3134, "bswaby@fbchtn.org", "Ben Swaby - FBCHville", "Test", messageDescription)

            #TODO rewrite message displayed
            # print '<p><h2>A charge of ${1} has been made </h2></p>'.format(model.Data.PaymentType, model.Data.PayAmount,model.Data.PaymentDescription)
            
        #TODO redirect back to charge
print '''<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
        </br>
        <p><h2>All participants of {2} have been charged.</h2></p>
        <a href="{0}/PyScript/MM-Charge?ProgramName={1}&ProgramID={2}"><i class="fa fa-home fa-3x"></i></a>'''.format(model.CmsHost, ProgramName, ProgramID,model.Data.pid)
