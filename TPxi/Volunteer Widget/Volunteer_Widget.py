#####################################################################
####WIDGET INFORMATION
#####################################################################
#This is a dashboard widget to show upcoming volunteer positions or
#link to your serve page.  Curently this is meant to work with 
#scheduler report.  You can find this report on my github here.
#https://github.com/bswaby/Touchpoint/tree/main/Blue%20Toolbar/Scheduler%20Report

# Written By: Ben Swaby (TPxi Software, LLC)
# Email: bswaby@fbchtn.org                                                                                                      
# Website: https://tpxisoftware.com
# GitHub: https://github.com/bswaby/Touchpoint  (50+ free tools)                                                                
# ----------------------------------------------------------------                                                              
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to                                                                
# support continued development, check out:                                                                                     
#
# DisplayCache(TM) - church digital signage that integrates with TouchPoint(R)                                                  
# https://displaycache.com                                
#
# TPxi Go(TM) - your church contacts, wherever you work.
# Look up anyone in TouchPoint(R), log calls and emails from Outlook                                                            
# or your phone. No tab switching, no lost context.
# https://tpxigo.com                                                                                                            
# ----------------------------------------------------------------

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
serveOpportunitiesLink = 'https://fbchville.com/serve/' 

#This is the Name of your scheduler list 
SchedulerReportPythonScriptName = 'ScheduleList' 

#####################################################################
####START OF CODE.  No configuration should be needed beyond this point
#####################################################################


sqlVolunteer = '''
    WITH OrgData AS (
        -- Fetch PeopleId, OrganizationId, and OrganizationName
        SELECT 
            om.PeopleId, 
            om.OrganizationId,
            o.OrganizationName  -- Add OrganizationName
        FROM OrganizationMembers om 
        LEFT JOIN Organizations o ON o.OrganizationId = om.OrganizationId
        WHERE om.PeopleId = {0}
        AND o.RegistrationTypeId = 22
    )
    
    , AllServices AS (
        SELECT 
            tsm.MeetingDateTime AS ServiceDate,
            FORMAT(tsm.MeetingDateTime, 'hh:mm tt') AS ServiceTime,
            FORMAT(tsm.MeetingDateTime, 'M/d/yy h:mm tt') AS ServiceDateTime,
            tstSG.NumberVolunteersNeeded AS Needed,
            tstSG.Require AS [Required],
            tsmt.TimeSlotMeetingTeamId,
            tsgrpvol.TimeSlotMeetingTeamSubGroupId,
            tssg.NumberVolunteersNeeded AS SubGroupsNeeded,
            p.Name2,
            p.Name,
            CASE 
                WHEN mt.Name IS NULL THEN tsmt.TeamName 
                ELSE mt.Name 
            END AS SubGroupTeam,
            tsmt.TeamName,
            od.OrganizationName,  -- Include OrganizationName in the results
    		od.OrganizationId  -- Incluce OrgId in results
        FROM TimeSlotMeetingTeams tsmt 
        LEFT JOIN TimeSlotMeetings tsm ON tsmt.TimeSlotMeetingId = tsm.TimeSlotMeetingId
        LEFT JOIN TimeSlotMeetingTeamSubGroups tssg ON tssg.TimeSlotMeetingTeamId = tsmt.TimeSlotMeetingTeamId
        LEFT JOIN TimeSlotTeamSubGroups tstSG ON tstSG.TimeSlotTeamSubGroupId = tssg.TimeSlotTeamSubGroupId
        LEFT JOIN Meetings m ON tsm.MeetingId = m.MeetingId
        LEFT JOIN TimeSlotMeetingTeamSubGroupVolunteers tsGrpVol 
            ON (tsGrpVol.IsActive <> 0 AND tsGrpVol.TimeSlotMeetingTeamId = tsmt.TimeSlotMeetingTeamId AND tsGrpVol.TimeSlotMeetingTeamSubGroupId IS NULL)
            OR (tsGrpVol.IsActive <> 0 AND tsGrpVol.TimeSlotMeetingTeamSubGroupId = tssg.TimeSlotMeetingTeamSubGroupId AND tsGrpVol.TimeSlotMeetingTeamSubGroupId IS NOT NULL)
        LEFT JOIN MemberTags mt ON mt.Id = tssg.MemberTagId
        LEFT JOIN People p ON p.PeopleId = tsGrpVol.PeopleId
        INNER JOIN OrgData od ON p.PeopleId = od.PeopleId
        WHERE 
            tsm.MeetingDateTime > GETDATE() 
            AND tsm.MeetingDateTime < DATEADD(DAY, 365, GETDATE()) 
            AND p.PeopleId = od.PeopleId
            AND m.OrganizationId = od.OrganizationId  
    )
    
    , AllPeeps AS (
        SELECT 
            STRING_AGG(a.Name2, ', ') AS Name2,
            a.ServiceDate,
            a.ServiceTime,
            a.ServiceDateTime,
            a.Needed,
            a.TimeSlotMeetingTeamId,
            a.SubGroupTeam,
            a.TeamName,
            a.Required,
    		a.OrganizationName,
    		a.OrganizationId
        FROM AllServices a
        GROUP BY 
            a.ServiceDate, a.ServiceTime, a.ServiceDateTime, a.Needed, 
            a.TimeSlotMeetingTeamId, a.SubGroupTeam, a.TeamName, a.Required, a.OrganizationName, a.OrganizationId
    )
    
    , ServiceCount AS (
        SELECT 
            ServiceDate,
            ServiceDateTime,
            ServiceTime,
            Needed,
            [Required], 
            TimeSlotMeetingTeamId,
            SubGroupTeam,
            TeamName,
            COUNT(Name2) AS Serving 
        FROM AllPeeps 
        GROUP BY 
            TimeSlotMeetingTeamId, TeamName, SubGroupTeam, ServiceDate, 
            ServiceDateTime, ServiceTime, Needed, [Required]  
    )
    
    , UniqueDates AS (
        SELECT DISTINCT
            DATEADD(DAY, DATEDIFF(DAY, 0, ServiceDateTime), 0) AS DateOnly,
            COUNT(DATEADD(DAY, DATEDIFF(DAY, 0, ServiceDateTime), 0)) AS CountDate 
        FROM ServiceCount
        GROUP BY DATEADD(DAY, DATEDIFF(DAY, 0, ServiceDateTime), 0)
    )
    
    SELECT
        sc.ServiceDate AS 'Date',
        sc.ServiceDateTime,
        FORMAT(sc.ServiceDate, 'dddd, MMMM dd, yyyy') AS DateOnly,
        ud.CountDate,
        CONVERT(CHAR(5), sc.ServiceDate, 108) AS [TimeOnly],
        a.TeamName AS [Team],
        sc.SubGroupTeam,
        CONVERT(VARCHAR(10), sc.Serving) + ' of ' + CONVERT(VARCHAR(10), sc.Needed) AS 'Filled',
        sc.Serving,
        sc.Needed,
        sc.Required,
        a.TimeSlotMeetingTeamId,
        ISNULL(
            SUBSTRING(
                (SELECT ', ' + a1.Name2 AS [text()]
                 FROM AllPeeps a1
                 WHERE a1.ServiceDateTime = sc.ServiceDateTime
                 AND a1.SubGroupTeam = sc.SubGroupTeam 
                 ORDER BY a1.Name2
                 FOR XML PATH ('')), 2, 1000),
            '--->none<---'
        ) AS Names,
    	a.OrganizationName,
    	a.OrganizationId
    FROM ServiceCount sc
    LEFT JOIN AllPeeps a ON a.TeamName = sc.TeamName 
        AND a.ServiceDateTime = sc.ServiceDateTime 
        AND a.SubGroupTeam = sc.SubGroupTeam 
    LEFT JOIN UniqueDates ud ON ud.DateOnly = DATEADD(DAY, DATEDIFF(DAY, 0, sc.ServiceDate), 0)
    WHERE sc.SubGroupTeam NOT LIKE 'EMS' 
    AND sc.SubGroupTeam NOT LIKE 'LEO'
    GROUP BY sc.ServiceDateTime, a.TeamName, sc.SubGroupTeam, sc.ServiceDate, sc.Serving, sc.Needed, a.TimeSlotMeetingTeamId, sc.Required, ud.CountDate, a.Name2, a.OrganizationName, a.OrganizationId
    ORDER BY sc.ServiceDate;


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
    </style>
'''

rsqlVolunteer = q.QuerySql(sqlVolunteer.format(model.UserPeopleId))  # 2832 Execute query
rsqlvVolunteerCount = 0
rsqlVolunteerCount = 0
rsqlOrgName = ''

if not rsqlVolunteer:
    print '<h4>Serve Opportunities</h4>'
else:
    #print '<h4>Upcoming Volunteer Schedule</h4>'
    
    #get initial count of row... probably a better way of doing this, but this works for now
    for vv in rsqlVolunteer:
        rsqlvVolunteerCount += 1
        
    for v in rsqlVolunteer:
        
        rsqlVolunteerCount += 1 #increment to help with writing out
        
        #need to write something to handle subgroups
        #if v.SubGroupTeam != '' and v.SubGroupTeam != v.Team:
        
        #Write section ending to start new name
        if rsqlOrgName != '' and v.OrganizationName != rsqlOrgName:
            print '</table></div></div>'
        
        #Write header for new v.OrganizationName
        if rsqlOrgName == '' or v.OrganizationName != rsqlOrgName:
            print '''<div class="timeline-item"><div class="timeline-date"><h4>Serve | {0}</h4></div><div class="timeline-content"><table style="border-collapse: collapse; width: 100%;">'''.format(v.OrganizationName)
    
        rsqlOrgName = v.OrganizationName
        
        #loop through all entires for v.OrganizationName
        
        #&#128336; 
        print '''
            <tr>
                <td style="padding: 5px; text-align: left;">📅 {0} {1}</td>
                <td style="padding: 5px; text-align: left;">📍 <a href="/PyScript/{4}?CurrentOrgId={3}" target="_blank">{2}</a></td>
            </tr>
            '''.format(v.DateOnly,v.TimeOnly,v.Team,v.OrganizationId,SchedulerReportPythonScriptName)
        
        if rsqlvVolunteerCount == 1 or rsqlVolunteerCount == rsqlvVolunteerCount:
            print '</table></div></div>'
            

print '''<a href="{}" target="_blank">Find your place to serve today.</a>'''.format(serveOpportunitiesLink)
print '<hr>'
