#####################################################################
####WIDGET INFORMATION
#####################################################################
#This is a dashboard widget to show upcoming mission trips for the
#the person viewing and their family.  A few features
# - Mission Trip Name
# - Other Family Members Going on the Trip
# - Status of Payment
# - Quick Link to the registration tab
# - Quick Link to the funding page
# - Quick Link to funding email
# - Link to Mission Page if they are leader
# - Link to Mission Page

#Note:This is meant to work with the Mission Page to allow the leaders
# to access a page of all those going on the trip.  You can find this
# code here. https://github.com/bswaby/Touchpoint/blob/main/Missions/MissionsDashboard/Missions%20Dashboard.py

#Adding the widget
#
#Step 1:  Add Python
#Paste this file in as a python script under Admin~Advanced~Special Content~Python
#Name the file whatever you like and make sure to add the word widget
#under content keywords
#
#Step 2: Add Widget
#Go to Admin~Advanced~HomePage Widgets
#Click add widget
#- Add Name, Description, Select Role(s)
#- Select the widget under Code (Python)
#Click Save and Then Enable the Widget

#####################################################################
####USER CONFIG FIELDS
#####################################################################

#Link to your serve page if they are not volunteering
serveOpportunitiesLink = 'http://fbchville.com/missions' 

#Name of Mission Dashboard Python Script
MissionPageLink = 'Mission_Dashboard'

#####################################################################
####START OF CODE.  No configuration should be needed beyond this point
#####################################################################


sqlMission = '''
    DECLARE @InvId INT;
    SET @InvId = {0}; -- 3089 4415 Example PersonId
    
    WITH FamilyMembers AS (
        -- Get FamilyId of @InvId
        SELECT FamilyId
        FROM People
        WHERE PeopleId = @InvId
    ),
    OrgMemberships AS (
        -- Get all OrganizationIds where @InvId is a member
        SELECT DISTINCT om.OrganizationId
        FROM OrganizationMembers om
        WHERE om.PeopleId = @InvId
    )
    SELECT  
        pro.Name AS [Program],
        pro.Id AS [ProgramId],
        d.Name AS [Division],
        d.Id AS [DivisionId],
        o.OrganizationName,
        o.OrganizationId,
        p.PeopleId,
        p.Name2,
        p.Age,
        p.FirstName,
        p.LastName,
        p.EmailAddress,
        p.CellPhone,
        p.HomePhone,
        p.FamilyId,
        COALESCE(SUM(ts.TotPaid), 0) AS [Paid],       -- Ensure transactions are optional
        COALESCE(FORMAT(SUM(ts.TotCoupon), 'C', 'en-US'), '$0.00') AS [Coupons],
        COALESCE(SUM(ts.IndDue), 0) AS [Outstanding],
        MAX(ts.TranDate) AS TranDate, -- Use MAX to avoid nulls
        COUNT(*) OVER () AS TotalResults,  -- Adds total count of results
		at.Description AS ClassRole
    FROM OrganizationMembers om
    INNER JOIN OrgMemberships orgm ON om.OrganizationId = orgm.OrganizationId -- Only include organizations of @InvId
    INNER JOIN People p ON om.PeopleId = p.PeopleId  
    INNER JOIN FamilyMembers fm ON p.FamilyId = fm.FamilyId -- Only include family members
	LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
	LEFT JOIN lookup.AttendType at ON at.Id = mt.AttendanceTypeId
    LEFT JOIN Organizations o ON o.OrganizationId = om.OrganizationId
    LEFT JOIN Division d ON d.Id = o.DivisionId
    LEFT JOIN Program pro ON pro.Id = d.ProgId
    LEFT JOIN TransactionSummary ts ON ts.PeopleId = p.PeopleId 
        AND ts.OrganizationId = om.OrganizationId  -- Keep transactions if they exist
        AND ts.IsLatestTransaction = 1
    WHERE 
        o.IsMissionTrip = 1
        AND o.OrganizationStatusId = 30  
    GROUP BY  
        d.Name, o.OrganizationName, o.OrganizationId, pro.Name, pro.Id, 
        d.Name, d.Id, p.PeopleId, p.Name2, p.Age, p.FirstName, p.LastName, 
        p.EmailAddress, p.CellPhone, p.HomePhone, p.FamilyId, at.Description
    ORDER BY o.OrganizationName, p.Name2;
'''

#model.UserPeopleId

print '''
    <style>
        .timeline {
            border-left: 4px solid #003366;
            padding-left: 20px;
        }
        
        .timeline-item {
            margin-bottom: 15px;
            position: relative;
        }
        
        .timeline-date {
            font-weight: bold;
            color: #003366;
        }
        
        .timeline-content {
            background: #f8f9fa;
            padding: 10px;
            border-radius: 5px;
        }
        .progress-container {
            width: 100%;
            max-width: 225px; /* Adjust width as needed */
            background-color: #eee;
            border-radius: 10px;
            overflow: hidden;
        }

        .progress-bar {
            width: calc(225 / 1800 * 100%); /* Adjust based on progress */
            height: 10px;
            background-color: #757575;
        }

        .progress-text {
            font-weight: bold;
            margin-bottom: 5px;
        }
    </style>
'''
#model.UserPeopleId

rsqlMission = q.QuerySql(sqlMission.format(model.UserPeopleId))  # 33467 Execute query  model.UserPeopleId
countLoops = 0
rsqlOrgName = ''

if not rsqlMission:
    print '<h4>Mission Opportunities</h4>'
    
else:
    print '''<div class="timeline-item"><div class="timeline-date"><h3>Missions</h3></div><div class="timeline-content">'''

    for v in rsqlMission:
    
        countLoops += 1
        NameTitle = ''
        
        # Format the funding progression text
        totalFee = v.Paid + v.Outstanding 
        
        if v.Paid >= totalFee:
            progress_text = "$%.2f of $%.2f <strong>goal met!</strong>" % (v.Paid, totalFee)
        else:
            progress_text = "$%.2f raised of <strong>$%.2f</strong> goal" % (v.Paid, totalFee)

        if v.Paid == 0 and totalFee == 0:
            progress_percentage = 100.00
            progress_text = "$%.2f <strong>fees charged!</strong>" % (totalFee)
        else:
            progress_percentage = (v.Paid / float(totalFee)) * 100  # Ensure float division
   
        #Set new Org Name
        if countLoops == 1 or rsqlOrgName != v.OrganizationName:
            TripTitle = '''<h4>{0}</h4>'''.format(v.OrganizationName)
        else:
            TripTitle = ''
            
        if v.TotalResults > 1:
            NameTitle = '''{0} {1}'''.format(v.FirstName,v.LastName)
            
        if v.ClassRole == 'Leader':
            LeaderLink = '''<a href="/PyScript/{2}?OrgView={0}"> ({1})</a>'''.format(v.OrganizationId,v.ClassRole,MissionPageLink)
            NameTitle += LeaderLink
        else:
            LeaderLink = ''

        rsqlOrgName = v.OrganizationName
            
        print '''
            {3}{6}
            <table style="border-collapse: collapse; width: 100%;">
            <tr>
                <td style="padding: 5px; text-align: left;"><a href="/Person2/{1}#tab-registrations">{5}</a>
                    <div class="progress-container">
                        <div class="progress-bar" style="width: {4:.2f}%;"></div>
                    </div><br>

                </td>
                <td style="padding: 5px; text-align: left;">&#127760; <a href="/OnlineReg/{2}/Giving/{1}">Funding</a></td>
                <td style="padding: 5px; text-align: left;">&#128231; <a href="/MissionTripEmail2/{2}/{1}">Supporters</a></td>
            </tr>
            </table>
            '''.format(v.Outstanding,v.PeopleId,v.OrganizationId,TripTitle,progress_percentage,progress_text,NameTitle)

print '''</div></div>
         <a href="{}" target="_blank">Explore ways you can be a part of missions.</a>
         <hr>
    '''.format(serveOpportunitiesLink)

#
#
