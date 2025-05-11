#####################################################################
####REPORT INFORMATION | Missions Dashbaord
#####################################################################

#The purpose of Missions Dashboard is meant to help give insight to 
#the mission leaders of active missions and outstanding payments

#-----------------This is highly configured for our environment------------------#
#----Considerable effort would need to be made to make it for your environment----#

#written by: Ben Swaby
#email: bswaby@fbchtn.org

#####################################################################
####START OF CODE.  No configuration should be needed beyond this point
#####################################################################
import datetime
import re

model.Header = "Missions Dashboard"

#style="border-top: 1px dashed lightblue;"


# Load the function definitions from the config
ConfigFile = '_FunctionLibrary'
FunctionLibrary = model.TextContent(ConfigFile)  # Loads function definitions as a string

# Normalize line endings to prevent syntax errors
FunctionLibrary = FunctionLibrary.replace('\r', '')  # Remove Windows-style carriage returns

# Execute the cleaned function definitions
exec(FunctionLibrary)  # This defines generate_html_table()


def progressbar_styles():
    return """
        <style>
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
        </style>"""
        
def popup_styles():
    return """
            <style>
            @media (max-width: 768px) {
                .email-popup {
                    width: 90%;
                    height: 90%;
                    top: 5%;
                    left: 5%;
                    transform: none;
                }
            }
            
            .email-popup {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: white;
                padding: 20px;
                border: 1px solid #ccc;
                box-shadow: 0px 0px 10px #aaa;
                max-width: 80%;
                max-height: 80%;
                overflow: auto;
                z-index: 1000;
                width: 600px; /* Default width for desktop */
                height: auto;
            }
            
            /* Close button */
            .close {
                color: red;
                float: right;
                font-size: 20px;
                cursor: pointer;
            }
        </style>
        
        <script>
            function showEmailBody(id) {
                document.getElementById(id).style.display = 'block';
            }
        
            function hideEmailBody(id) {
                document.getElementById(id).style.display = 'none';
            }
        </script>
    """
        
def dashboard_styles():
    return """
            <hr>
            <style>
            .chart-container {
                display: flex;
                justify-content: left;
                gap: 20px;
                margin-top: 20px;
            }
            .chart-box {
                display: flex;
                flex-direction: column;
                align-items: center;  
                text-align: center;
            }
            .value-text {
                font-size: 18px;
                font-weight: bold;
                margin-top: -10px;
            }

             * {
              box-sizing: border-box;
            }
            
            .container {
              display: flex; 
              align-items: flex-start; 
              gap: 0px; 
              padding: 0;
              margin: 0;
            }
            
            .parent-container {
              padding: 0;
              margin: 0;
              width: 100%;
            }
            .maincolumna {
              justify-content: left;
              background-color: white;
              flex: 4; /* Takes 80% of space */
              padding: 0; /* Ensure no extra padding */
              margin: 0;  /* Remove default margins */
              width: 100%; /* Ensures it stretches fully */
              box-sizing: border-box; /* Ensures padding doesn't affect width */
            }

            .no-padding td {
                line-height: 1 !important;
                padding: 2px !important;  /* Small padding */
                line-height: 1.5 !important;  /* Adjust line-height for text spacing */
                padding-left: 15px !important;
            }
            .rightcolumna {
              background-color: white;
              flex: 2; /* Takes 20% of space */
              padding: 10px;
              margin-top: 0px;
              text-align: left;
            }
            
            @media only screen and (max-width: 620px) {
              .container {
                flex-direction: column; /* Stack them in smaller screens */
              }
              .maincolumna, .rightcolumna {
                width: 100%;
              }
            }
            div[data-v-30ee886b] {
                display: none !important;
            }
            .container-fluid {
                background-color: white !important;
            }
            .box.box-responsive {
                    border: none !important;
            }
            td, th {
              padding: 8px;
              font-size: 14px; /* Default size */
            }
            /* Set a specific width for a particular <td> */
            td.bgcheck {
              width: 60px; /* Example fixed width at 14px font-size */
            }
            td.title {
              white-space: nowrap;  /* Prevent text from wrapping */
              overflow: hidden;      /* Hide overflowing text */
              text-overflow: ellipsis; /* Add "..." for overflow */
              max-width: 200;      /* Set a fixed or max width (required for ellipsis to work) */
              padding: 2px;
            }
            
            td.title h4 {
              display: inline-block; /* Ensure it respects width */
              max-width: 100%; /* Prevent it from expanding beyond td */
              overflow: hidden;
              text-overflow: ellipsis;
              white-space: nowrap;
              max-width: 200px;
              padding: 2px;
            }
            /* When the screen width is 600px or smaller, reduce text size */
            @media (max-width: 800px) {
              td, th {
                font-size: 12px;
                padding: 4px; /* Reduce padding to save space */
              }
              /* Adjust width for the specific <td> when font size is 12px */
              td.bgcheck {
                width: 50px; /* Example smaller width at 12px font-size */
              }
              td.title {
                white-space: nowrap;  /* Prevent text from wrapping */
                overflow: hidden;      /* Hide overflowing text */
                text-overflow: ellipsis; /* Add "..." for overflow */
                max-width: 150px;      /* Set a fixed or max width (required for ellipsis to work) */
                padding: 2px;
              }
              td.title h4 {
                display: inline-block; /* Ensure it respects width */
                max-width: 100%; /* Prevent it from expanding beyond td */
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                max-width: 150px;
                padding: 2px;
              }
              .maincolumna {
                padding: 0 !important; /* Override any padding */
                margin: 0 !important;
              }
            }

            .styled-table {
                width: 100%;
                max-width: 600px;
                border-collapse: collapse;
                background: white;
                border-radius: 8px;
                box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.1);
                overflow: hidden;
                font-size: 14px; /* Matches your default font */
            }
            
            .styled-table th, .styled-table td {
                padding: 8px;
                text-align: left;
            }
            
            .styled-table th {
                background-color: #3498db;
                color: white;
                font-size: 14px;
                text-transform: uppercase;
            }
            
            .styled-table tr:hover {
                background-color: #f1f1f1;
            }
            
            .styled-table .category {
                font-weight: bold;
                background-color: #ecf0f1;
            }
            
            /* Responsive Table Adjustments */
            @media (max-width: 800px) {
                .styled-table {
                    font-size: 12px;
                }
                .styled-table th, .styled-table td {
                    padding: 6px;
                }
            }
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
            
            /* simple Menu style */

            .simplenav-menu {
                display: flex;
                gap: 7px; /* Reduced spacing between menu items */
                padding: 10px;
            }
            .simplenav-menu a {
                color: black;
                text-decoration: none;
                padding: 8px 12px; /* Slightly adjusted padding */
                border-radius: 8px;
                transition: color 0.3s ease, background 0.3s ease;
            }
            .simplenav-menu a:hover {
                background: #f0f0f0;
            }
            /* Highlight only the active tab with background */
            .simplenav-menu a.simplenav-active {
                color: #4A90E2; /* Blue text for active tab */
                font-weight: bold;
                background: rgba(74, 144, 226, 0.1); /* Light blue highlight */
                padding: 8px 12px;
            }
            /* Custom Tooltip styles to pop-up explanation */
            .custom-tooltip {
                display: none;
                position: absolute;
                background-color: rgba(0, 0, 0, 0.7);
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
                max-width: 200px;
                z-index: 1000;
                box-shadow: 0 0 10px rgba(0, 0, 0, 0.2);
            }
    
            /* Custom Icon style */
            .custom-icon {
                font-size: 30px;
                cursor: pointer;
            }
        </style>

        <!-- Navigation Menu -->
        <div class="simplenav-menu">
            <a href="?simplenav=dashboard">Dashboard</a>
            <a href="?simplenav=due">Finance</a>
            <a href="?simplenav=messages">Messages</a>
            <a href="?simplenav=stats">Stats</a>
        </div>
    
        <script>
            // Get the URL parameter
            const urlParams = new URLSearchParams(window.location.search);
            const activeTab = urlParams.get("simplenav") || "dashboard"; // Default to Discover
    
            // Select all menu links
            const menuLinks = document.querySelectorAll(".simplenav-menu a");
    
            // Loop through links and add 'simplenav-active' class to the current tab
            menuLinks.forEach(link => {
                if (link.getAttribute("href") === `?simplenav=${activeTab}`) {
                    link.classList.add("simplenav-active");
                }
            });
        
            function toggleCustomTooltip(event, tooltipId) {
                event.stopPropagation();  // Prevent the click from propagating to the document
            
                var tooltip = document.getElementById(tooltipId);
                
                // Check if the tooltip is visible or not and toggle accordingly
                if (tooltip.style.display === 'none' || tooltip.style.display === '') {
                    // Position the tooltip near the icon
                    tooltip.style.top = event.pageY + 10 + 'px';
                    tooltip.style.left = event.pageX + 10 + 'px';
                    tooltip.style.display = 'block';  // Show the tooltip
                } else {
                    tooltip.style.display = 'none';  // Hide the tooltip
                }
            }
            
            // Hide all tooltips when clicking outside
            document.addEventListener('click', function(event) {
                // Find all tooltips and hide them
                var tooltips = document.querySelectorAll('.custom-tooltip');
                tooltips.forEach(function(tooltip) {
                    tooltip.style.display = 'none';
                });
            });

        </script>
    """

#Default View w/no options
if (model.Data.OrgView == '' or model.Data.OrgView is None) and model.Data.simplenav != 'stats' and model.Data.simplenav != 'due' and model.Data.simplenav != 'messages':
    
    sqlTotalMemberOutstanding = '''
        SELECT  
        ISNULL((Select Sum(Due) AS IndDue 
          From MissionTripTotals MTT
            LEFT JOIN Organizations O ON MTT.InvolvementId = O.OrganizationId
          WHERE o.IsMissionTrip = 1 AND o.OrganizationStatusId = 30 and mtt.name <> 'total'), 0) AS TotalOutstanding,
          
        ISNULL((SELECT SUM(o.MemberCount) 
        FROM Organizations o  
        INNER JOIN OrganizationExtra oe -- Changed to INNER JOIN since we only want records with a Close date
            ON oe.OrganizationId = o.OrganizationId 
            AND oe.Field = 'Close'
            AND oe.DateValue > GETDATE() -- Only include organizations with future close dates
        WHERE o.IsMissionTrip = 1  
          AND o.OrganizationStatusId = 30
          AND o.OrganizationId NOT IN (2736,2737,2738)), 0) AS TotalMembers,
          
        ISNULL((SELECT SUM(o.MemberCount) 
        FROM Organizations o  
        WHERE o.OrganizationId IN (2736,2737,2738)), 0) AS TotalApplications
    '''

    

    sqlActiveMissions = '''
        -- ============================
        -- Precompute EventStatus in a Temp Table
        -- ============================
        SELECT 
            oe.OrganizationId,
            MAX(CASE WHEN oe.Field = 'Main Event Start' THEN oe.DateValue END) AS StartDate,
            MAX(CASE WHEN oe.Field = 'Main Event End' THEN oe.DateValue END) AS EndDate,
            MAX(CASE WHEN oe.Field = 'Close' THEN oe.DateValue END) AS Closed,
            CASE 
                WHEN MAX(CASE WHEN oe.Field = 'Main Event Start' THEN oe.DateValue END) > GETDATE() THEN 'Pre'
                WHEN MAX(CASE WHEN oe.Field = 'Main Event Start' THEN oe.DateValue END) <= GETDATE() 
                     AND MAX(CASE WHEN oe.Field = 'Main Event End' THEN oe.DateValue END) >= GETDATE() THEN 'Event'
                WHEN MAX(CASE WHEN oe.Field = 'Close' THEN oe.DateValue END) < GETDATE() THEN 'Closed'
                WHEN MAX(CASE WHEN oe.Field = 'Main Event End' THEN oe.DateValue END) < GETDATE() THEN 'Post'
                ELSE 'Unknown'
            END AS EventStatus
        INTO #EventStatus
        FROM OrganizationExtra oe
        INNER JOIN Organizations o ON o.OrganizationId = oe.OrganizationId
        WHERE o.IsMissionTrip = 1 
        AND o.OrganizationStatusId = 30
        GROUP BY oe.OrganizationId;
        
        -- ============================
        -- Precompute MeetingCounts
        -- ============================
        SELECT 
            m.OrganizationId,
            COUNT(CASE WHEN m.MeetingDate < GETDATE() THEN 1 END) AS PastMeetings,
            COUNT(CASE WHEN m.MeetingDate >= GETDATE() THEN 1 END) AS FutureMeetings,
            STRING_AGG(
                CAST(
                    CASE 
                        WHEN m.MeetingDate < GETDATE() THEN 'Past - ' 
                        ELSE 'Upcoming - ' 
                    END 
                    + CONVERT(VARCHAR, m.MeetingDate, 120) + ' | ' 
                    + COALESCE(NULLIF(m.Description, ''), 'No Description') 
                AS NVARCHAR(MAX)), '<br>'
            ) AS MeetingList,
        	CASE 
        		WHEN o.RegStart > GETDATE() AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Pending'
        		WHEN o.RegStart IS NULL AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Open'
        		WHEN o.RegEnd < GETDATE() THEN 'Closed'
        		ELSE 'Open'
        	END AS RegistrationStatus
        INTO #MeetingCounts
        FROM Meetings m
        INNER JOIN Organizations o ON o.OrganizationId = m.OrganizationId
        WHERE o.IsMissionTrip = 1 
        AND o.OrganizationStatusId = 30
        GROUP BY m.OrganizationId,o.RegStart,o.RegEnd;
        
        -- ============================
        -- Precompute BackgroundCheckCounts
        -- ============================
        SELECT 
            om.OrganizationId,
            COUNT(CASE WHEN vs.Description = 'Complete' 
                       AND v.ProcessedDate >= DATEADD(YEAR, -3, GETDATE()) THEN 1 END) AS BackgroundCheckGood,
            COUNT(CASE WHEN vs.Description <> 'Complete' 
                       AND v.ProcessedDate >= DATEADD(YEAR, -3, GETDATE()) THEN 1 END) AS BackgroundCheckBad,
            COUNT(CASE WHEN v.ProcessedDate IS NULL 
                       OR v.ProcessedDate < DATEADD(YEAR, -3, GETDATE()) THEN 1 END) AS BackgroundCheckMissing,
            STRING_AGG(
                CAST(
                    CASE 
                        WHEN vs.Description <> 'Complete' 
                        AND v.ProcessedDate >= DATEADD(YEAR, -3, GETDATE()) THEN 'Bad - ' + p.Name + ' (' + CAST(p.Age AS VARCHAR(3)) + ')'
                        WHEN v.ProcessedDate IS NULL OR v.ProcessedDate < DATEADD(YEAR, -3, GETDATE()) THEN 'Missing - ' + p.Name + ' (' + CAST(p.Age AS VARCHAR(3)) + ')'
                        ELSE NULL 
                    END 
                AS NVARCHAR(MAX)), '<br>'
            ) AS PeopleWithBadOrMissingBackgroundChecks
        INTO #BackgroundCheckCounts
        FROM OrganizationMembers om
        LEFT JOIN Volunteer v ON v.PeopleId = om.PeopleId
        LEFT JOIN lookup.VolApplicationStatus vs ON vs.Id = v.StatusId
        LEFT JOIN People p ON p.PeopleId = om.PeopleId
        INNER JOIN Organizations o ON o.OrganizationId = om.OrganizationId
        WHERE o.IsMissionTrip = 1 
        AND o.OrganizationStatusId = 30
        AND om.MemberTypeId <> 230
        GROUP BY om.OrganizationId;
        
        -- ============================
        -- Precompute PassportCounts
        -- ============================
        SELECT 
            om.OrganizationId,
            COUNT(DISTINCT om.PeopleId) AS TotalPeople,
            COUNT(DISTINCT CASE WHEN r.passportnumber IS NOT NULL AND r.passportexpires IS NOT NULL THEN om.PeopleId END) AS PeopleWithPassports,
            STRING_AGG(CAST(CASE WHEN r.passportnumber IS NULL OR r.passportexpires IS NULL THEN p.Name ELSE NULL END AS NVARCHAR(MAX)), '<br>') AS PeopleMissingPassports
        INTO #PassportCounts
        FROM OrganizationMembers om
        INNER JOIN recreg r ON om.PeopleId = r.PeopleId
        INNER JOIN People p ON p.PeopleId = om.PeopleId
        INNER JOIN Organizations o ON o.OrganizationId = om.OrganizationId
        WHERE o.IsMissionTrip = 1 
        AND o.OrganizationStatusId = 30
        AND om.MemberTypeId <> 230
        GROUP BY om.OrganizationId;
        
        -- ============================
        -- Precompute MemberDetails
        -- ============================
        SELECT 
            om.OrganizationId,
            STRING_AGG(p.Name + ' (' + CAST(p.Age AS VARCHAR(3)) + ') - ' + mt.Description, '<br>') AS PeopleList
        INTO #MemberDetails
        FROM OrganizationMembers om
        INNER JOIN People p ON p.PeopleId = om.PeopleId
        LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
        INNER JOIN Organizations o ON o.OrganizationId = om.OrganizationId
        WHERE o.IsMissionTrip = 1 
        AND o.OrganizationStatusId = 30
        AND om.MemberTypeId <> 230
        GROUP BY om.OrganizationId;
        
        -- ============================
        -- Final Query with Joins
        -- ============================
        SELECT  
            od.OrganizationId,
            --od.Program,
            od.OrganizationName, 
            od.MemberCount,
            od.ImageUrl,
            ISNULL(SUM(mtt.Due), 0) AS Outstanding,
            ISNULL(SUM(mtt.Raised), 0) AS TotalDue,
        	ISNULL(SUM(mtt.TripCost), 0) AS TotalFee,
            MAX(COALESCE(mc.PastMeetings, 0)) AS PastMeetings,
            MAX(COALESCE(mc.FutureMeetings, 0)) AS FutureMeetings,
            MAX(COALESCE(bc.BackgroundCheckGood, 0)) AS BackgroundCheckGood,
            MAX(COALESCE(bc.BackgroundCheckBad, 0)) AS BackgroundCheckBad,
            MAX(COALESCE(bc.BackgroundCheckMissing, 0)) AS BackgroundCheckMissing,
            MAX(COALESCE(pc.PeopleWithPassports, 0)) AS PeopleWithPassports,
            md.PeopleList,
            mc.MeetingList,
            pc.PeopleMissingPassports,
            bc.PeopleWithBadOrMissingBackgroundChecks,
            es.StartDate,
            es.EndDate,
            es.Closed,
            COALESCE(es.EventStatus, 'Dates Not Set') AS EventStatus,
        	COALESCE(mc.RegistrationStatus, 'Reg Dates Not Set') AS RegistrationStatus
        FROM Organizations od
        LEFT JOIN #EventStatus es ON od.OrganizationId = es.OrganizationId
        LEFT JOIN #MeetingCounts mc ON od.OrganizationId = mc.OrganizationId
        LEFT JOIN #BackgroundCheckCounts bc ON od.OrganizationId = bc.OrganizationId
        LEFT JOIN #PassportCounts pc ON od.OrganizationId = pc.OrganizationId
        LEFT JOIN #MemberDetails md ON od.OrganizationId = md.OrganizationId
        --LEFT JOIN TransactionSummary ts ON od.OrganizationId = ts.OrganizationId
        LEFT JOIN MissionTripTotals mtt ON mtt.InvolvementId = od.OrganizationId
        WHERE od.IsMissionTrip = 1 AND od.OrganizationStatusId = 30 and mtt.name <> 'total'
            {0}
        GROUP BY 
            od.OrganizationId,
            --od.Program,
            od.OrganizationName,
            od.MemberCount,
            od.ImageUrl,
            es.StartDate,
            es.EndDate,
            es.Closed,
            es.EventStatus,
        	mc.RegistrationStatus,
            mc.MeetingList,
            md.PeopleList,
            pc.PeopleMissingPassports,
            bc.PeopleWithBadOrMissingBackgroundChecks

        Order By es.EventStatus Desc,
        	es.StartDate,
        	od.OrganizationName;
        
        -- Cleanup temp tables
        DROP TABLE #EventStatus;
        DROP TABLE #MeetingCounts;
        DROP TABLE #BackgroundCheckCounts;
        DROP TABLE #PassportCounts;
        DROP TABLE #MemberDetails;
    '''
    
    sqlActiveMissionLeaders = '''
        Select p.PeopleId
        	,p.Name
        	,p.Age
        	,om.OrganizationId
        	,mt.Description AS Leader
        From OrganizationMembers om
            LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
            LEFT JOIN lookup.AttendType at ON at.Id = mt.AttendanceTypeId
            LEFT JOIN People p ON p.PeopleId = om.PeopleId
        Where OrganizationId = {0} and at.id = 10 
    '''

    sqlDashboardMeetings = '''
        SELECT 
            o.OrganizationName,
            m.Description,
            m.Location,
            FORMAT(m.MeetingDate, 'dddd, MMMM dd h:mm tt') AS FormattedMeetingDate, 
            m.OrganizationId
        FROM Meetings m
        LEFT JOIN Organizations o ON m.OrganizationId = o.OrganizationId
        LEFT JOIN Division d ON d.Id = o.DivisionId  
        LEFT JOIN Program pro ON pro.Id = d.ProgId  
        WHERE 
            o.IsMissionTrip = 1  
            AND o.OrganizationStatusId = 30  
            AND m.MeetingDate >= CAST(GETDATE() AS DATE)  -- Include today and future dates
        ORDER BY m.MeetingDate;
        '''
    sqlPaymentStatus = '''
      Select 
    	  Sum(TripCost) AS TotalFee, 
    	  Sum(Raised) AS Paid, 
    	  Sum(Due) AS Outstanding 
      From MissionTripTotals MTT
      LEFT JOIN Organizations O ON MTT.InvolvementId = O.OrganizationId
      WHERE o.IsMissionTrip = 1 AND o.OrganizationStatusId = 30 
    '''
    sqlPaymentStatusold = '''
        SELECT         
            -- Aggregate transaction data to avoid duplication
            ISNULL(sum(ts_data.TotalPaid), 0) AS [Paid],
            ISNULL(sum(ts_data.TotalOutstanding), 0) AS [Outstanding],
            ISNULL(sum(ts_data.TotalFee), 0) AS [TotalFee]
            
        FROM Organizations o
        LEFT JOIN OrganizationMembers om ON o.OrganizationId = om.OrganizationId
        LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
        INNER JOIN People p ON om.PeopleId = p.PeopleId
    
        
        -- Pre-aggregate transaction data per person to avoid duplication
        OUTER APPLY (
            SELECT 
                SUM(ts.TotPaid) AS TotalPaid,
                SUM(ts.TotCoupon) AS TotalCoupons,
                SUM(ts.IndDue) AS TotalOutstanding,
                SUM(ts.TotalFee) AS TotalFee
            FROM TransactionSummary ts
            WHERE ts.OrganizationId = o.OrganizationId
            AND ts.PeopleId = p.PeopleId  -- Ensuring transactions match the person
            AND IsLatestTransaction = 1 
        ) ts_data
        
        WHERE o.IsMissionTrip = 1 AND o.OrganizationStatusId = 30 -- Ensure you're filtering for a specific org
    
    '''
     
    sqlLast20Signups = '''
        SELECT Top 20
            o.OrganizationId,
            o.OrganizationName, 
    		p.Name,
    		pic.SmallUrl,
    		om.EnrollmentDate
        FROM Organizations o  
    		INNER JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
    		INNER JOIN People p ON p.PeopleId = om.PeopleId
    		LEFT JOIN Picture pic ON pic.PictureId = p.PictureId
        WHERE o.IsMissionTrip = 1  
            AND OrganizationStatusId = 30  
            AND om.MemberTypeId <> 230
    	ORDER BY om.EnrollmentDate Desc
    '''
        
    TotalMembers = 0
    TotalOutstanding = 0
    rsqlTotalMemberOutstanding = q.QuerySql(sqlTotalMemberOutstanding)
    for tmo in rsqlTotalMemberOutstanding:
        TotalMembers = tmo.TotalMembers
        TotalOutstanding = tmo.TotalOutstanding   
        TotalApplications = tmo.TotalApplications
    
    print dashboard_styles()
    print popup_styles()
    
    print '''
        <hr>
        <div style="display: flex; gap: 20px; justify-content: flex-start; width: 100%; margin-left: 0;">
            <!-- People Card -->
            <div style="background: #f8fafc; border-radius: 10px; padding: 15px 20px; text-align: center; 
                        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); min-width: 150px;">
                <div style="font-size: 24px; font-weight: bold; color: #333;">''' + str(TotalMembers) + '''</div>
                <div style="font-size: 14px; color: #555;">Active</div>
            </div>
            <div style="background: #f8fafc; border-radius: 10px; padding: 15px 20px; text-align: center; 
                        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); min-width: 150px;">
                <div style="font-size: 24px; font-weight: bold; color: #333;">''' + str(TotalApplications) + '''</div>
                <div style="font-size: 14px; color: #555;">In the Queue</div>
            </div>
        
            <!-- Amount Due Card -->
            <div style="background: #f8fafc; border-radius: 10px; padding: 15px 20px; text-align: center; 
                        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); min-width: 150px;">
                <div style="font-size: 24px; font-weight: bold; color: #d9534f;">''' + format_currency(TotalOutstanding, True, True) + '''</div>
                <div style="font-size: 14px; color: #555;">Due</div>
            </div>
        </div>
        '''
    print progressbar_styles()
        
    print '''
        <!-- Tooltip explanations -->
        <div id="calendar-tooltip" class="custom-tooltip">Meetings: Upcoming-Completed</div>
        <div id="passport-tooltip" class="custom-tooltip">Passports: Active</div>
        <div id="bgcheck-tooltip" class="custom-tooltip">Background Check: Approved-Failed-None</div>
        <div id="registered-tooltip" class="custom-tooltip">Total in the Involvement</div>
        
        <div class="container">
        <div class="maincolumna">                                                                                                           
            <p>
            <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 800px; width: 100%; overflow-x: auto;">
    '''
    
    #simple url link to show / hide closed
    if model.Data.ShowClosed == '1':
        ShowClosed = """ AND es.EventStatus = 'Closed'"""
        print """<a href="?ShowClosed=0">Hide Closed</a>"""
    else:
        ShowClosed = """ AND (es.EventStatus is NULL or es.EventStatus != 'Closed')"""
        print """<a href="?ShowClosed=1">Show Closed</a>"""


    rsqlActiveMissions = q.QuerySql(sqlActiveMissions.format(ShowClosed))
    #sqlActiveMissionLeaders

    last_registration_status = None  # Initialize a variable to track the previous RegistrationStatus
    
    for am in rsqlActiveMissions:
        # Check if the current RegistrationStatus is different from the last one
        if am.EventStatus != last_registration_status:
            # If it differs, print a new row with the RegistrationStatus as a separator
            print '''
                <tr style="border-top: 1px dashed lightblue;">
                    <td colspan="1" style="text-align: center; padding: 10px; font-weight: bold; background-color: #f0f0f0;">
                        {0}
                    </td>
                    <td data-label="Meetings" style="padding: 2px; text-align: center; vertical-align: middle; background-color: #f0f0f0;">
                        <span id="calendar-icon" class="custom-icon" onclick="toggleCustomTooltip(event, 'calendar-tooltip')">     
                            <h4>&#x1F4C5;</h4>
                        </span>
                    </td>
                    <td data-label="Passport" style="padding: 2px; text-align: center; vertical-align: middle; background-color: #f0f0f0;">
                        <span id="passport-icon" class="custom-icon" onclick="toggleCustomTooltip(event, 'passport-tooltip')">
                            <h4>&#x1F194;</h4>
                        </span>
                    </td>
                    <td class="bgcheck" data-label="BG Check" style="padding: 2px; text-align: center; vertical-align: middle; text-align: center; vertical-align: middle; background-color: #f0f0f0;">
                        <span id="bgcheck-icon" class="custom-icon" onclick="toggleCustomTooltip(event, 'bgcheck-tooltip')">    
                            <h4>&#x1F6E1;</h4>
                        </span>
                    </td>
                    <td data-label="Attending" style="padding: 2px; text-align: center; vertical-align: middle; background-color: #f0f0f0;">
                        <span id="registered-icon" class="custom-icon" onclick="toggleCustomTooltip(event, 'registered-tooltip')">
                            <span style="font-size: 28px;"><h4>&#x1F465;</h4></span>
                        </span>
                    </td>
                </tr>
            '''.format(str(am.EventStatus))
            
            # Update the last_registration_status to the current one
            last_registration_status = am.EventStatus
    
        # Main row for the mission data
        if am.ImageUrl != None or am.ImageUrl == '':
            imgUrl = '&#x1F3DE;'
        else:
            imgUrl = ''
        
        MeetingList = ''
        MeetingList = '- no meetings - ' if not am.MeetingList else MeetingList + am.MeetingList
        
        print '''
            <tr style="border-top: 1px solid lightblue;">
                <td style="padding: 2px;">{7}<a href="?OrgView={1}"><b>{2}</b></a></td>
                <td style="padding: 2px; text-align: center; vertical-align: middle;">
                    <a href="#" onclick="showEmailBody('meetings_{2}'); return false;">{6}-{5}</a>
                    <div id="meetings_{2}" class="email-popup" style="display: none;">
                        <div class="email-content">
                            <span class="close" onclick="hideEmailBody('meetings_{2}')">&times;</span>
                            <h3>Meetings</h3>
                            <div>{13}</div>
                        </div>
                    </div>  
                </td>
                <td style="padding: 2px; text-align: center; vertical-align: middle;">
                    <a href="#" onclick="showEmailBody('missingpassport_{2}'); return false;">{11}</a>
                    <div id="missingpassport_{2}" class="email-popup" style="display: none;">
                        <div class="email-content">
                            <span class="close" onclick="hideEmailBody('missingpassport_{2}')">&times;</span>
                            <h3>Missing Passport</h3>
                            <div>{14}</div>
                        </div>
                    </div>  
                </td>
                <td style="padding: 2px; text-align: center; vertical-align: middle;">
                    <a href="#" onclick="showEmailBody('bgcheckissue_{2}'); return false;">{8}-{9}-{10}</a>
                    <div id="bgcheckissue_{2}" class="email-popup" style="display: none;">
                        <div class="email-content">
                            <span class="close" onclick="hideEmailBody('bgcheckissue_{2}')">&times;</span>
                            <h3>Background Checks Issues</h3>
                            <div>{15}</div>
                        </div>
                    </div>  
                </td>
                <td style="padding: 2px; text-align: center; vertical-align: middle;">
                    <a href="#" onclick="showEmailBody('orgId_{2}'); return false;">{3}</a>
                    <div id="orgId_{2}" class="email-popup" style="display: none;">
                        <div class="email-content">
                            <span class="close" onclick="hideEmailBody('orgId_{2}')">&times;</span>
                            <h3>{2}</h3>
                            <div>{12}</div>
                        </div>
                    </div>                
                </td> 
            </tr>
            <tr class="no-padding">
                <td colspan="5" style="padding-left: 5px; font-style: italic; color: #666;">
                    <span style="background-color: AliceBlue;">Trip: {16} - {17}</span>
                </td>
            </tr>
        '''.format(
            str(am.RegistrationStatus),
            str(am.OrganizationId),
            am.OrganizationName,
            str(am.MemberCount),
            format_currency(am.Outstanding, True, True),
            am.PastMeetings,
            am.FutureMeetings,
            imgUrl,
            am.BackgroundCheckGood,
            am.BackgroundCheckBad,
            am.BackgroundCheckMissing,
            am.PeopleWithPassports,
            am.PeopleList,
            MeetingList,
            am.PeopleMissingPassports,
            am.PeopleWithBadOrMissingBackgroundChecks,
            format_date(str(am.StartDate), 'short'),
            format_date(str(am.EndDate), 'short')
        )

        # Process active mission leaders
        rsqlActiveMissionLeaders = q.QuerySql(sqlActiveMissionLeaders.format(str(am.OrganizationId)))
        if not rsqlActiveMissionLeaders:
            print '''
                <tr class="no-padding">
                    <td colspan="5" style="padding-left: 5px; font-style: italic; color: #666;">
                        <span style="background-color: yellow;">-- No leader(s) --</span>
                    </td>
                </tr>
            '''
        else:
            for aml in rsqlActiveMissionLeaders:
                print '''
                    <tr class="no-padding">
                        <td colspan="3" style="padding-left: 5px;">
                            <a href="/Person2/{2}" target="_blank"><i class="fa fa-info-circle"></i></a>&nbsp;<i>{0}: {1} ({3})</i>
                        </td>
                        <td></td>
                        <td></td>
                    </tr>
                '''.format(aml.Leader, aml.Name, aml.PeopleId, aml.Age)

        
        if am.TotalFee:
            if am.TotalDue >= am.TotalFee:
                progress_text = "$%.2f of $%.2f <strong>goal met!</strong>" % (am.TotalDue, am.TotalFee)
            else:
                progress_text = "<strong>$%.2f remaining</strong> of $%.2f goal" % (am.Outstanding, am.TotalFee)
    
            if am.TotalDue == 0 and am.TotalFee == 0:
                progress_percentage = 100.00
                progress_text = "$%.2f <strong>fees charged!</strong>" % (am.TotalFee)
            else:
                progress_percentage = (am.TotalDue / float(am.TotalFee)) * 100  # Ensure float division
            
            print """
                <tr class="no-padding">
                    <td colspan="3" style="padding-left: 5px;">
                        <p class="progress-text">
                            <div class="progress-container">
                                <div class="progress-bar" style="width: {1:.2f}%;"></div>
                            </div>{0}
                        </p>
                    </td>
                    <td></td>
                    <td></td>
                </tr>
                """.format(progress_text,progress_percentage)


    print '''</table>
        <h3>Last 20 Sign-ups</h3>
        <p>
            <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 800px; width: 100%; overflow-x: auto;">
        '''
    rsqlLast20Signups = q.QuerySql(sqlLast20Signups) 
    for ls in rsqlLast20Signups:
        print '''
                    <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                        <td data-label="name" style="padding: 2px;">{1}</td>
                        <td data-label="enrollmentdate" style="padding: 2px;">{2}</td>
                        <td data-label="orgname" style="padding: 2px;">{3}</td>
                    </tr>
    
            '''.format(ls.SmallUrl,
                       ls.Name,
                       ls.EnrollmentDate,
                       ls.OrganizationName,
                       ls.OrganizationId)

      
        
    print '''</table></p></div>
        <div class="rightcolumna">
            <h3>Upcoming Meetings</h3>
            <p>'''
    
    rsqlDashboardMeetings = q.QuerySql(sqlDashboardMeetings) 
    if not rsqlDashboardMeetings:
        print """
            <p>Looks like the ministry calendars are as clear as a sunny day! 🌞</p>
            <p>Why not schedule some meetings and make it a bit more exciting?</p>
        """

    for rdm in rsqlDashboardMeetings:
        print '''<div class="timeline-item">
                    <div class="timeline-date"><a href="/PyScript/Mission_Dashboard?OrgView={4}" target="_blank">{2}</a></div>
                    <div class="timeline-content">
                        <h2 style="margin: 0; padding: 0; font-size: 20px;">{3}</h2>
                        <p>⏰ {0} 📍 {1}</p>
                    </div>
                </div>
                <hr style="margin-top: 0; margin-bottom: 0; padding: 0;">
            </p>'''.format(rdm.FormattedMeetingDate,rdm.Location,rdm.OrganizationName,rdm.Description,rdm.OrganizationId)
    
    print '''</div></div>'''

if model.Data.simplenav == 'due':

    sqlOutstandingPayments = '''
        Select 
          --Sum(TripCost) AS TotalFee, 
          p.PeopleId,
          o.OrganizationId,
          O.OrganizationName,
          p.Name2,
          Sum(mtt.Raised) AS Paid, 
          Sum(mtt.Due) AS Outstanding 
        From MissionTripTotals MTT
        LEFT JOIN Organizations O ON MTT.InvolvementId = O.OrganizationId
        LEFT JOIN People p ON p.PeopleId = MTT.PeopleId
        WHERE o.IsMissionTrip = 1 AND o.OrganizationStatusId = 30 AND mtt.Due <> 0  and mtt.name <> 'total'
        GROUP BY  p.Name2, o.OrganizationId ,O.OrganizationName,p.PeopleId
        ORDER BY o.organizationName, p.Name2
    '''

    sqlOutstandingPaymentsold = '''
        Select 
         pro.Name AS [Program]
         ,pro.Id AS [ProgramId]
         ,d.Name AS [Division]
         ,d.Id AS [DivisionId]
         ,o.OrganizationName
         ,o.OrganizationId
         ,ts.PeopleId
         ,p.Name2
         ,p.Age
         ,p.FirstName
         ,p.LastName
         ,p.EmailAddress
         ,p.CellPhone
         ,p.HomePhone
         ,p.FamilyId
         ,FORMAT(Sum(ts.TotPaid), 'C', 'en-US') AS [Paid]
         ,FORMAT(Sum(ts.TotCoupon), 'C', 'en-US') AS [Coupons]
         ,FORMAT(Sum(ts.IndDue), 'C', 'en-US') AS [Outstanding]
         ,ts.TranDate
        FROM [TransactionSummary] ts
        INNER JOIN [People] p on ts.PeopleId = p.PeopleId
        LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
        LEFT JOIN Division d ON d.Id = o.DivisionId
        LEFT JOIN Program pro ON pro.Id = d.ProgId
        Where 
         ts.IndDue <> 0
         AND ts.IsLatestTransaction = 1
         AND o.IsMissionTrip = 1
         AND o.OrganizationStatusId = 30  
        Group By 
         d.Name
         ,o.OrganizationName
         ,o.OrganizationId
         ,pro.Name
         ,pro.Id
         ,d.Name
         ,d.Id
         ,ts.PeopleId
         ,p.Name2
         ,p.Age
         ,p.FirstName
         ,p.LastName
         ,p.EmailAddress
         ,p.CellPhone
         ,p.HomePhone
         ,p.FamilyId
         ,ts.TranDate
        Order by o.OrganizationName,p.Name2
        '''
        
    sqlTransactions = '''
        SELECT TOP 100 
        	tl.TransactionDate,
        	tl.People as Name,
        	o.OrganizationName,
        	tl.BegBal,
        	tl.TotalPayment,
        	tl.TotDue,
        	tl.Message
        
        FROM TransactionList tl
        	LEFT JOIN Organizations o ON o.OrganizationId = tl.OrgId
        WHERE
        	o.IsMissionTrip = 1  
            AND o.OrganizationStatusId = 30
        	AND tl.BegBal is not null
        ORDER BY tl.Id DESC
    '''
    
    print dashboard_styles()
    
    print '''
        <hr>
        <h3>Due</h3>
        <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 800px; width: 100%;">
            <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                <td style="padding: 2px;"></td>
                <td style="padding: 2px;"><h4>Paid</h4></td>
                <td style="padding: 2px;"><h4>Due</h4></td>
            </tr>
        '''
    
    rsqlOutstandingPayments = q.QuerySql(sqlOutstandingPayments)
    
    last_org_name = None  # Track last organization to insert separator rows
 
    for op in rsqlOutstandingPayments:
         # Add a separator row when OrganizationName changes
        if op.OrganizationName != last_org_name:
            print '''
                <tr>
                    <td colspan="3" style="background: lightgray; font-weight: bold; padding: 5px;">{0}</td>
                </tr>
            '''.format(op.OrganizationName)
            last_org_name = op.OrganizationName  # Update tracke
 
 
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;"><a href="?OrgView={4}#APeopleId{5}">{1}</a></td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{2}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{3}</td>
            </tr>
            '''.format(op.OrganizationName
                      ,op.Name2
                      ,format_currency(op.Paid,True,True)
                      ,format_currency(op.Outstanding,True,True)
                      ,op.OrganizationId
                      ,op.PeopleId)
    
    print '''</table>'''               

    print '''
        <hr>
        <h3>Last 100 Transactions</h3>
        <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
            <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%; ">
                <td style="padding: 2px;border-top: 1px dashed #D6EBFF;"><h4>Date</h4></td>
                <td style="padding: 2px;border-top: 1px dashed #D6EBFF;"><h4>Name</h4></td>
                <td style="padding: 2px;border-top: 1px dashed #D6EBFF;min-width: 40px; max-width: 60px; text-align: right; white-space: nowrap;"><h4>Beg. Bal.</h4></td>
                <td style="padding: 2px;border-top: 1px dashed #D6EBFF;min-width: 40px; max-width: 60px; text-align: right; white-space: nowrap;"><h4>Pmt</h4></td>
                <td style="padding: 2px;border-top: 1px dashed #D6EBFF;min-width: 40px; max-width: 60px; text-align: right; white-space: nowrap;"><h4>Due</h4></td>
            </tr>
        '''

    rsqlTransactions = q.QuerySql(sqlTransactions)
    
    for t in rsqlTransactions:
        print '''
            <tr style="border-top: 1px solid lightblue;">
                <td style="padding: 2px;"><span style="color: black; font-weight: 350;">{0}</span></td>
                <td style="padding: 2px;"><span style="color: black; font-weight: 350;">{1}</span></td>
                <td style="padding: 2px;min-width: 40px; max-width: 60px; text-align: right; white-space: nowrap;"><span style="color: black; font-weight: 350;">{2}</span></td>
                <td style="padding: 2px;min-width: 40px; max-width: 60px; text-align: right; white-space: nowrap;"><span style="color: black; font-weight: 350;">{3}</span></td>
                <td style="padding: 2px;min-width: 40px; max-width: 60px; text-align: right; white-space: nowrap;"><span style="color: black; font-weight: 350;">{4}</span></td>
            </tr>
            '''.format(t.TransactionDate,
                       t.Name,
                       format_currency(t.BegBal, True, True),
                       format_currency(t.TotalPayment, True, True),
                       format_currency(t.TotDue, True, True))
        
        #background-color: #f9f9f9;
        print '''
            <tr style="border-top: 1px dashed #D6EBFF;">
                <td style="padding: 2px; padding-left: 6px;border-right: 1px dashed lightblue;" colspan="2"><span style="color: black; opacity: 0.5;">{0}</span></td>
                <td style="padding: 2px; text-align: left;" colspan="3"><span style="color: black; opacity: 0.5;">{1}</span></td>
            </tr>
            '''.format(t.OrganizationName,
                       t.Message)
            
    print '''</table>'''
    
    

if model.Data.simplenav == 'stats':
    
    sqlMembership = '''
        Select  ms.Description AS MemberStatus
            ,Count(Distinct p.PeopleId) AS [Count]
        From Organizations o
            LEFT JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
            LEFT JOIN People p ON p.PeopleId = om.PeopleId
            LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
        Where o.IsMissionTrip = 1
            AND om.InactiveDate is NULL
            AND o.OrganizationStatusId = 30  
        Group By ms.Description
        HAVING COUNT(DISTINCT p.PeopleId) > 0  
        Order by ms.Description
    '''
    
    sqlBaptism = '''
        Select 
             bs.Description AS [BaptismStatus]
            ,Count(Distinct p.PeopleId) AS [Count]
        From Organizations o
            LEFT JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
            LEFT JOIN People p ON p.PeopleId = om.PeopleId
            LEFT JOIN lookup.BaptismStatus bs ON bs.Id = p.BaptismStatusId
        Where 
            o.IsMissionTrip = 1
            AND o.OrganizationStatusId = 30  
        Group By bs.Description
        HAVING COUNT(DISTINCT p.PeopleId) > 0  
        Order by bs.Description
        '''
        
    sqlAgeBin = '''
        SELECT  
            FLOOR(p.Age / 10) * 10 AS AgeBin,  -- Group ages into 10-year bins (e.g., 20-29, 30-39, etc.)
            COUNT(DISTINCT p.PeopleId) AS [Count]
        FROM Organizations o
            LEFT JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
            LEFT JOIN People p ON p.PeopleId = om.PeopleId
        WHERE o.IsMissionTrip = 1
          AND o.OrganizationStatusId = 30
          AND p.Age IS NOT NULL  -- Exclude NULL ages
        GROUP BY FLOOR(p.Age / 10) * 10  -- Grouping by the 10-year age bin
        ORDER BY AgeBin;
        '''
    
    sqlGender = '''
        Select 
        	g.Description as [Gender]
        	,Count(Distinct p.PeopleId) AS [Count] 
        From Organizations o
            LEFT JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
            LEFT JOIN People p ON p.PeopleId = om.PeopleId
            LEFT JOIN lookup.Gender g ON g.Id = p.GenderId
        Where o.IsMissionTrip = 1
            AND o.OrganizationStatusId = 30  
            AND g.Description is Not NULL
        Group By g.Description
    '''

    print dashboard_styles()

    print'''<hr>
            <div class="rightcolumna"><p>
                </p>
                    <h3>Stats</h3>
                    <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 300px; width: 100%;">
                        <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                            <td style="padding: 2px;"><h4>Membership</h4></td>
                            <td style="padding: 2px;"></td>
                        </tr>
                '''
            
    rsqlMembership = q.QuerySql(sqlMembership)
    for m in rsqlMembership:
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{0}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{1}</td>
            </tr>
            '''.format(m.MemberStatus,m.Count)

    
    print '''
        </table>
        <br>
        <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 300px; width: 100%;">
            <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                <td style="padding: 2px;"><h4>Baptism Status</h4></td>
                <td style="padding: 2px;"></td>
            </tr>
        '''
    
    rsqlBaptism = q.QuerySql(sqlBaptism)
    for b in rsqlBaptism:
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{0}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{1}</td>
            </tr>
            '''.format(b.BaptismStatus,b.Count)
    
    print '''</table>
       <br>
        <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 300px; width: 100%;">
            <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                <td style="padding: 2px;"><h4>Gender</h4></td>
                <td style="padding: 2px;"></td>
            </tr>
        '''
    
    rsqlGender = q.QuerySql(sqlGender)
    for g in rsqlGender:
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{0}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{1}</td>
            </tr>
            '''.format(g.Gender,g.Count)
    
    print '''</table>
        <br>
        <table class="styled-table" id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 300px; width: 100%;">
            <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                <td style="padding: 2px;"><h4>Age Bin</h4></td>
                <td style="padding: 2px;"></td>
            </tr>'''
    rsqlAgeBin = q.QuerySql(sqlAgeBin)
    for ab in rsqlAgeBin:
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{0}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{1}</td>
            </tr>
            '''.format(ab.AgeBin,ab.Count)
    
    print '''</table></p></div></div>'''

if model.Data.OrgView != '':
    sqlInvolvementInformation = '''
    WITH OrgData AS (
        SELECT  
            o.OrganizationId,
            pro.Name AS [Program],
            pro.Id AS [ProgramId],
            o.OrganizationName, 
    		o.Description,
    		o.ImageUrl,
            -- Handle NULL values for AccountingCode and DonationFundId
            ISNULL(o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'INT'), 0) AS AccountingCode,
            ISNULL(o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'INT'), 0) AS DonationFundId,
            ISNULL(o.RegSettingXML.value('(/Settings/Fees/Fee)[1]', 'FLOAT'), 0) AS Fee,  -- Handle NULL
            ISNULL(o.RegSettingXML.value('(/Settings/Fees/Deposit)[1]', 'FLOAT'), 0) AS Deposit, -- Handle NULL
            o.RegStart,
            o.RegEnd,
            -- Determine Registration Status
            CASE 
                WHEN o.RegStart > GETDATE() AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Pending'
                WHEN o.RegStart IS NULL AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Active'
                WHEN o.RegEnd < GETDATE() THEN 'Closed'
                ELSE 'Active'  -- Default to Active if no conditions are met
            END AS RegistrationStatus
        FROM Organizations o  
            LEFT JOIN Division d ON d.Id = o.DivisionId  
            LEFT JOIN Program pro ON pro.Id = d.ProgId  
        WHERE o.IsMissionTrip = 1  
            AND OrganizationStatusId = 30  
    		AND OrganizationId = {0}
    )
    SELECT  
        od.OrganizationId,
        od.Program,
        od.ProgramId,
        od.OrganizationName, 
    	od.Description,
    	od.ImageUrl,
        od.AccountingCode,  -- Now guaranteed to be 0 instead of NULL
        od.DonationFundId,   -- Now guaranteed to be 0 instead of NULL
        FORMAT(od.Fee, 'C', 'en-US') AS Fee,  -- Formatting after handling NULL
        FORMAT(od.Deposit, 'C', 'en-US') AS Deposit,
        FORMAT(ISNULL(SUM(ts.IndDue), 0), 'C', 'en-US') AS Outstanding,  -- Handle NULL before formatting
        od.RegistrationStatus
    FROM OrgData od  
        LEFT JOIN TransactionSummary ts ON od.OrganizationId = ts.OrganizationId  
    GROUP BY  
        od.OrganizationId,
        od.Program,
        od.ProgramId,
        od.OrganizationName,
    	od.Description,
    	od.ImageUrl,
        od.AccountingCode,
        od.DonationFundId,
        od.Fee,
        od.Deposit,
        od.RegistrationStatus
    ORDER BY od.OrganizationId;
    '''
    sqlInvolvementMembers = '''
    WITH BackgroundCheckCounts AS (
        SELECT 
            om.OrganizationId,
            om.PeopleId,
            -- Checking if 'Complete' status is within the last 3 years, otherwise 'No'
            CASE 
                WHEN vs.Description = 'Complete' 
                AND v.ProcessedDate >= DATEADD(YEAR, -3, GETDATE()) 
                THEN 'Yes'
                ELSE 'No'
            END AS BackgroundCheckStatus
        FROM OrganizationMembers om
        LEFT JOIN Volunteer v ON v.PeopleId = om.PeopleId
        LEFT JOIN lookup.VolApplicationStatus vs ON vs.Id = v.StatusId
        WHERE om.OrganizationId = {0}  AND om.MemberTypeId <> 230 -- Apply the WHERE condition here as well
        GROUP BY om.OrganizationId, om.PeopleId, vs.Description, v.ProcessedDate
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
        bs.Description AS BaptismStatus,
        mt.Description AS OrgMemType,
        ms.Description AS MemberStatus,
        pic.SmallUrl AS Picture,
        rr.MedicalDescription,
        rr.emcontact,
        rr.emphone,
        
        -- Check if both passport fields have data
        CASE 
            WHEN rr.passportnumber IS NOT NULL AND rr.passportexpires IS NOT NULL THEN 'Yes'
            ELSE 'No'
        END AS HasPassport,
    
        -- BackgroundCheckStatus is now a 'Yes' or 'No' based on the logic in BackgroundCheckCounts
        bc.BackgroundCheckStatus AS BackgroundCheckStatus,
        
        -- Aggregate transaction data to avoid duplication
        ISNULL(ts_data.TotalPaid, 0) AS [Paid],
        --FORMAT(ISNULL(ts_data.TotalCoupons, 0), 'C', 'en-US') AS [Coupons],
        FORMAT(ISNULL(ts_data.TotalOutstanding, 0), 'C', 'en-US') AS [Outstanding],
        ISNULL(ts_data.TotalFee, 0) AS [TotalFee]
        
    FROM Organizations o
    LEFT JOIN OrganizationMembers om ON o.OrganizationId = om.OrganizationId
    LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
    INNER JOIN People p ON om.PeopleId = p.PeopleId
    LEFT JOIN Division d ON d.Id = o.DivisionId
    LEFT JOIN Program pro ON pro.Id = d.ProgId
    LEFT JOIN Picture pic ON pic.PictureId = p.PictureId
    LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
    LEFT JOIN lookup.BaptismStatus bs ON bs.Id = p.BaptismStatusId
    LEFT JOIN RecReg rr ON rr.PeopleId = p.PeopleId
    
    -- Join with the BackgroundCheckCounts CTE
    LEFT JOIN BackgroundCheckCounts bc ON bc.OrganizationId = o.OrganizationId
    AND bc.PeopleId = p.PeopleId  -- Join on both OrganizationId and PeopleId
    
    -- Pre-aggregate transaction data per person to avoid duplication
    OUTER APPLY (
        /**SELECT 
            ts.PeopleId,
            SUM(ts.TotPaid) AS TotalPaid,
            SUM(ts.TotCoupon) AS TotalCoupons,
            SUM(ts.IndDue) AS TotalOutstanding,
            SUM(ts.TotalFee) AS TotalFee
        FROM TransactionSummary ts
        WHERE ts.OrganizationId = o.OrganizationId
        AND ts.PeopleId = p.PeopleId  -- Ensuring transactions match the person
        AND IsLatestTransaction = 1 
        GROUP BY ts.PeopleId**/
        Select MTT.PeopleId, 
            Sum(MTT.TripCost) AS TotalFee, 
            Sum(MTT.Raised) AS TotalPaid, 
            Sum(MTT.Due) AS TotalOutstanding 
        From MissionTripTotals MTT
        WHERE MTT.InvolvementId = o.OrganizationId
            AND MTT.PeopleId = p.PeopleId
         Group By PeopleId
    ) ts_data
    
    WHERE o.OrganizationId = {0} AND om.MemberTypeId <> 230 -- Ensure you're filtering for a specific org
    ORDER BY mt.Description, p.Name2;
    '''

    
    sqlDocs = '''
        SELECT rmt.Description AS ResourceType,rt.Name
        	,rc.Name as ResourceGroup
        	,r.Name as Resource
        	,r.ResourceDescription
			,r.Description
        	,r.ResourceId
			,r.ResourceUrl
			,CASE WHEN r.Description is Not Null Then r.Description ELSE r.ResourceUrl END AS Content
			,r.*
			,rt.*
        FROM [Resource] r
            LEFT JOIN ResourceCategory rc ON rc.ResourceCategoryId = r.ResourceCategoryId
            LEFT JOIN ResourceType rt ON rt.ResourceTypeId = r.ResourceTypeId
            LEFT JOIN ResourceOrganization ro ON ro.ResourceId = r.ResourceId
			LEFT JOIN Lookup.ResourceMediaType rmt ON rmt.Id = r.ResourceMediaTypeId
        Where rt.Name = 'Missions'
            AND (VisibleToEveryone = 1 or ro.OrganizationId = {0})
        Order by rt.Name, rc.Name, r.DisplayOrder
    '''
    
    
    sqlDocAttachements = '''
        SELECT Name
              ,FilePath 
        FROM ResourceAttachment
        WHERE ResourceId = {0}
        ORDER BY DisplayOrder
    '''
    

    #<img src="{0}" alt="Mission Image Missing" style="width: 100%; height: auto;">
    #<img src="{0}" onerror="this.onerror=null; this.src='https://c4265878.ssl.cf2.rackcdn.com/fbchville.2502091537.Hey__I_am_beautiful._Consider_adding_a_photo.png';" alt="Image">
    rsqlInvolvementInformation = q.QuerySql(sqlInvolvementInformation.format(str(model.Data.OrgView)))
    for ii in rsqlInvolvementInformation:
        print '''
            <img src="{0}" alt="Mission Image Missing" style="width: 100%; height: auto;">
            

            <hr>
        
            <h2><a href="/Org/{6}" target="_blank"><i class="fa-solid fa-people-roof"></i></a>&nbsp{1}</h2>
            <strong>Status:</strong> {2}<br>
            <strong>Total Outstanding:</strong> {3}<br>
            <strong>Registration Start:</strong> {4}<br>
            <strong>Registration End:</strong> {5}<br><hr>
            '''.format(ii.ImageUrl,ii.OrganizationName,ii.RegistrationStatus,str(ii.Outstanding),str(ii.RegStart),str(ii.RegEnd),str(ii.OrganizationId))
    
    #---------------------------------------------
    #Print out past and upcoming meetings
    #---------------------------------------------
    print get_meetings(str(model.Data.OrgView))
    
    last_group = None  # Track the last ResourceGroup to avoid repeating it
    
    print '<h2 style="margin: 0; color: #2c3e50;">Resources</h2>'
    last_group = None  # Track the last ResourceGroup to avoid repeating it
    resource = ''

    #-----------------------------------
    #Resources specific to Missions
    #-----------------------------------
    print get_resources('Missions',str(model.Data.OrgView),'Serve')

    
    print '''        
            <h3>Team</h3>
            <table id="trnoborder" style="border: 0px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1200px; width: 100%;">
                <tr id="trnoborder" style="border: 0px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1200px; width: 100%;">
                    <td style="padding: 2px;"><h4></h4></td>
                    <td style="padding: 2px;"><h4></h4></td>
                    <td style="padding: 2px;"><h4></h4></td>
                </tr>
            '''
    orgMemType = ''    
    rsqlInvolvementMembers = q.QuerySql(sqlInvolvementMembers.format(str(model.Data.OrgView)))
    for im in rsqlInvolvementMembers:
        formatted_homephone = '☎️ <a href="tel:' + im.HomePhone + '">' + model.FmtPhone(im.HomePhone) + '</a><br>' if im.HomePhone and im.HomePhone != im.CellPhone else ''
        

        # Format the progress text
        if im.Paid >= im.TotalFee:
            progress_text = "$%.2f of $%.2f <strong>goal met!</strong>" % (im.Paid, im.TotalFee)
        else:
            progress_text = "$%.2f raised of <strong>$%.2f</strong> goal" % (im.Paid, im.TotalFee)

        if im.Paid == 0 and im.TotalFee == 0:
            progress_percentage = 100.00
            progress_text = "$%.2f <strong>fees charged!</strong>" % (im.TotalFee)
        else:
            progress_percentage = (im.Paid / float(im.TotalFee)) * 100  # float division
        
        print progressbar_styles()
        print '''
            <style>
                /* ensure printing */
                @media print {
                    .progress-bar {
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                    .progress-container {
                        display: block !important; 
                    }
                
                    .progress-bar {
                        display: block !important;
                        background-color: #757575 !important; 
                    }
                }
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
        if orgMemType == '' or orgMemType != im.OrgMemType:
            orgMemType = im.OrgMemType
            print '''<tr style="background-color: #004085; color: white; font-size: 18px; font-weight: bold; text-align: center;">
                        <td colspan="5" style="padding: 8px;">
                            {0}
                        </td>
                    </tr>
                 
                    '''.format(im.OrgMemType)
        #<img src="{12}" alt="Need Picture">
        if im.HasPassport == 'Yes':
            passport = """<span>&#x1F194;</span>"""
        else:
            passport = """
                        <div style="display: inline-block; position: relative; font-size: inherit; line-height: 1;">
                            <span style="opacity: 0.5;">&#x1F194;</span>
                            <div style="
                                position: absolute;
                                top: 50%;
                                left: 0;
                                width: 100%;
                                height: 2px;  /* Slash thickness */
                                background-color: red;
                                transform: rotate(-45deg);
                            "></div>
                        </div>
                    `   """
        
        
        if im.BackgroundCheckStatus == 'Yes':
            BackgroundCheckStatus = """<span>&#x1F6E1;</span>"""
        else:
            BackgroundCheckStatus = """
                                    <div style="display: inline-block; position: relative; font-size: inherit; line-height: 1;">
                                        <span style="opacity: 0.5;">&#x1F6E1;</span>
                                        <div style="
                                            position: absolute;
                                            top: 50%;
                                            left: 0;
                                            width: 100%;
                                            height: 2px;  /* Slash thickness */
                                            background-color: red;
                                            transform: rotate(-45deg);
                                        "></div>
                                    </div>
                                    """
        
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">
                    <a href="/Person2/{10}" target="_blank"><i class="fa fa-info-circle"></i></a>
                    &nbsp<span id="APeopleId{10}">{1}</span> ({2})
                    <br>{17}{18}
                </td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">📱 {5}</td>
            </tr>
            <tr>
                <td>
                    <img src="{11}" onerror="this.onerror=null; this.src='https://c4265878.ssl.cf2.rackcdn.com/fbchville.2502091552.Hey__I_am_beautiful._Consider_adding_a_photo_-1-.png';" alt="Image">
                    <br>
                    <p class="progress-text">{15}
                    <div class="progress-container">
                        <div class="progress-bar" style="width: {16:.2f}%;"></div>
                    </div></p>
                </td> 
                <td colspan="1" style="padding: 2px; font-size: 14px; color: gray; vertical-align: top;">
                    {6}
                    {9}<br>
                    Baptism: {3}<br>
                    Church: {4}<hr style="margin: 0; padding: 0; border: none; border-top: 1px solid #999;">
                    Emergency Contact: {12}<br>
                    Emergency Phone: {13}
                </td>
            </tr>

            '''.format(im.OrgMemType 
                        ,im.Name2 
                        ,im.Age
                        ,im.BaptismStatus
                        ,im.MemberStatus
                        ,'<a href="tel:{}">{}</a>'.format(im.CellPhone, model.FmtPhone(im.CellPhone)) if im.CellPhone else ''
                        ,formatted_homephone
                        ,im.Outstanding
                        ,im.Paid
                        #,im.Coupons
                        ,'<a href="https://myfbch.com/Org/{0}#tab-Members-tab">{1}</a>'.format(im.OrganizationId, im.EmailAddress) if im.EmailAddress else ''
                        ,im.PeopleId
                        ,im.Picture
                        ,im.emcontact
                        ,'<a href="tel:{}">{}</a>'.format(im.emphone, model.FmtPhone(im.emphone)) if im.emphone else ''
                        ,im.OrganizationId
                        ,progress_text
                        ,progress_percentage
                        ,passport
                        ,BackgroundCheckStatus)
                        
            #&nbsp<a href="/OnlineReg/{15}/Giving/{11}" target="_blank"><i class="fa fa-cross"></i></a>
                        
            #<td style="border-top: 1px dashed lightblue; padding: 2px;">% Checks:</td>        
            #<td colspan="1" style="padding: 2px; font-size: 14px; color: gray; vertical-align: top;">
            #    Future Check 1:<br>
            #    Future Check 2:<br>
            #</td> 
                    
            #support link = https://myfbch.com/OnlineReg/2779/Giving/33467
            
            #                <!-- Contact Row (Indented slightly for readability) -->
            #    <tr>
            #        <td></td> <!-- Empty cell for alignment -->
            #        <td colspan="1" style="padding: 2px; font-size: 14px; color: gray;">
            #            📱 {5} | ☎️ {6} 
            #        </td>
            #    </tr>
            #| 📧 {10}
            
    print '''</table>'''

if model.Data.simplenav == 'messages':
    
    sqlEmails = '''
        SELECT 
            eqt.Id, 
            CAST(eqt.Sent AS DATE) AS SentDate,
            eq.Subject, 
            eq.Body, 
            COUNT(eqt.PeopleId) AS PeopleCount,
            STRING_AGG(p.Name, '<br>') WITHIN GROUP (ORDER BY p.Name2) AS PeopleNames,
            -- Add prefix to OrganizationName
            CASE 
                WHEN o.RegStart > GETDATE() AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Pending: '
                WHEN o.RegStart IS NULL AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Open: '
                WHEN o.RegEnd < GETDATE() THEN 'Closed: '
                ELSE 'Open: '
            END + o.OrganizationName AS OrganizationName,
            
            -- Sorting helper: Assign a numeric priority for sorting
            CASE 
                WHEN o.RegEnd < GETDATE() THEN 2  -- Closed should be sorted last
                WHEN o.RegStart > GETDATE() THEN 0  -- Pending first
                ELSE 1  -- Open in the middle
            END AS SortOrder
        
        FROM EmailQueueTo eqt
            LEFT JOIN EmailQueue eq ON eq.Id = eqt.Id
            LEFT JOIN People p ON p.PeopleId = eqt.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
        WHERE 
            o.IsMissionTrip = 1  
            AND o.OrganizationStatusId = 30
        GROUP BY 
            eqt.Id, eq.Subject, CAST(eqt.Sent AS DATE), eq.Body, o.OrganizationName, 
            o.RegStart, o.RegEnd
        {0}
        ORDER BY 
            SortOrder ASC,  -- Ensures "Closed" is sorted last
            OrganizationName ASC,  -- Sort alphabetically within each category
            eqt.Id;


    '''

      

    print dashboard_styles()
    print popup_styles()
    
    #simple url link to show / hide single values with typically represent individual testing or people signing up
    if model.Data.ShowSingle == '1':
        ShowSingle = ''
        print """<a href="?simplenav=messages">Hide Single Messages</a>"""
    else:
        ShowSingle = 'HAVING COUNT(eqt.PeopleId) > 1'
        print """<a href="?simplenav=messages&ShowSingle=1">Show Single Messages</a>"""

    
    print ''' 
        <table id="trnoborder" style="border: 0px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1200px; width: 100%;">

        '''
    
    rsqlEmails = q.QuerySql(sqlEmails.format(ShowSingle))
    last_org_name = None  # Track last organization to insert separator rows
    
    for e in rsqlEmails:
        # Check if e.SentDate is not None
        if e.SentDate:
            try:
                # Convert e.SentDate to datetime if it's not already
                sent_date = datetime.datetime.strptime(str(e.SentDate), '%m/%d/%Y %I:%M:%S %p')
                sent_date = sent_date.strftime('%m/%d/%Y')  # Format to only show the date
            except ValueError:
                sent_date = e.SentDate  # Fallback if parsing fails
        else:
            sent_date = ''
        
        # Convert email body safely and fix image paths
        email_body = e.Body.replace('"', '&quot;').replace("'", "&apos;")
        email_body = re.sub(r'src=["\'](?!https?:)(.*?)["\']', r'src="https://yourwebsite.com/\1"', email_body)
    
        # Add a separator row when OrganizationName changes
        if e.OrganizationName != last_org_name:
            print '''
                <tr>
                    <td colspan="4" style="background: lightgray; font-weight: bold; padding: 5px;">{0}</td>
                </tr>
            '''.format(e.OrganizationName)
            last_org_name = e.OrganizationName  # Update tracker
    
        # Print main email rows
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{0}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">
                    <a href="#" onclick="showEmailBody('email_{3}'); return false;">{1}</a>
                    <div id="email_{3}" class="email-popup" style="display: none;">
                        <div class="email-content">
                            <span class="close" onclick="hideEmailBody('email_{3}')">&times;</span>
                            <h3>{1}</h3>
                            <div>{4}</div>
                        </div>
                    </div>
                </td>
                <td style="border-top: 1px dashed lightblue; padding: 2px; text-align: center;">
                    <a href="#" onclick="showEmailBody('people_{3}'); return false;">{2}</a>
                    <div id="people_{3}" class="email-popup" style="display: none;">
                        <div class="email-content">
                            <span class="close" onclick="hideEmailBody('people_{3}')">&times;</span>
                            <h3>Sent To</h3>
                            <div>{5}</div>
                        </div>
                    </div>
                </td>
            </tr>
            '''.format(sent_date,
                       e.Subject,
                       e.PeopleCount,
                       e.Id,  # Unique ID for modals
                       email_body,
                       e.PeopleNames)


    
    print '</table>'
