#####################################################################
####REPORT INFORMATION | Missions Dashbaord
#####################################################################
#This is ##beta## and still being developed.

#The purpose of Missions Dashboard is meant to help give insight to 
#the mission leaders of active missions and outstanding payments



#####################################################################
####START OF CODE.  No configuration should be needed beyond this point
#####################################################################

model.Header = "Missions Dashboard"

#Default View w/no options
if model.Data.OrgView == '' or model.Data.OrgView is None:
    
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
    
    sqlTotalMemberOutstanding = '''
        SELECT  
        ISNULL((SELECT SUM(ts.IndDue) 
                       FROM TransactionSummary ts 
                       JOIN Organizations o ON ts.OrganizationId = o.OrganizationId 
                       WHERE o.IsMissionTrip = 1  
                       AND o.OrganizationStatusId = 30  
                       AND ts.IsLatestTransaction = 1), 0) AS TotalOutstanding,  -- Sum of outstanding payments formatted as currency
    
        ISNULL((SELECT SUM(o.MemberCount) 
                FROM Organizations o  
                WHERE o.IsMissionTrip = 1  
                AND o.OrganizationStatusId = 30), 0) AS TotalMembers  -- Sum of members without duplication
        '''
    
    sqlActiveMissions = '''
        WITH OrgData AS (
            SELECT  
                o.OrganizationId,
                pro.Name AS [Program],
                pro.Id AS [ProgramId],
                o.OrganizationName, 
                o.MemberCount,
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
                    WHEN o.RegStart IS NULL AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Open'
                    WHEN o.RegEnd < GETDATE() THEN 'Closed'
                    ELSE 'Open'  -- Default to Active if no conditions are met
                END AS RegistrationStatus
            FROM Organizations o  
                LEFT JOIN Division d ON d.Id = o.DivisionId  
                LEFT JOIN Program pro ON pro.Id = d.ProgId  
            WHERE o.IsMissionTrip = 1  
                AND OrganizationStatusId = 30  
                AND OrganizationId Not In (2737,2738)
        )
        SELECT  
            od.OrganizationId,
            od.Program,
            od.ProgramId,
            od.OrganizationName, 
            od.MemberCount,
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
            od.MemberCount,
            od.AccountingCode,
            od.DonationFundId,
            od.Fee,
            od.Deposit,
            od.RegistrationStatus
        ORDER BY od.OrganizationName,od.RegistrationStatus;
        '''
    
    sqlOutstandingPayments = '''
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
        
    TotalMembers = 0
    TotalOutstanding = 0
    rsqlTotalMemberOutstanding = q.QuerySql(sqlTotalMemberOutstanding)
    for tmo in rsqlTotalMemberOutstanding:
        TotalMembers = tmo.TotalMembers
        TotalOutstanding = tmo.TotalOutstanding
    
    print '''

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
                align-items: center;  /* Centers both the chart and text */
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
              display: flex; /* Enables Flexbox */
              align-items: flex-start; /* Aligns items to the top */
              gap: 20px; /* Adds spacing between columns */
            }
            
            .maincolumna {
              background-color: white;
              flex: 4; /* Takes 80% of space */
              padding: 0 10px;
            }
            
            .rightcolumna {
              background-color: white;
              flex: 1; /* Takes 20% of space */
              padding: 0px;
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
        </style>

        <div class="chart-container">
            <div class="chart-box">
                <div id="outstanding_chart"></div>
                <div id="outstanding_value" class="value-text">$0</div>  <!-- Display formatted value -->
            </div>
            <div class="chart-box">
                <div id="attending_chart"></div>
                <div class="value-text">65</div>  <!-- Static display -->
            </div>
        </div>

        <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
        <script type="text/javascript">
            google.charts.load('current', {packages: ['gauge']});
            google.charts.setOnLoadCallback(drawCharts);
    
            function drawCharts() {
                var outstandingData = google.visualization.arrayToDataTable([
                    ['Label', 'Value'],
                    ['Due', 0]  // Start at 0
                ]);
    
                var attendingData = google.visualization.arrayToDataTable([
                    ['Label', 'Value'],
                    ['Attending', 0]  // Start at 0
                ]);
    
                var options = {
                    width: 250, height: 150,
                    redFrom: 40000, redTo: 50000,
                    yellowFrom: 25000, yellowTo: 39999,
                    greenFrom: 0, greenTo: 24999,
                    minorTicks: 5,
                    max: 50000,  // Adjusted max range
                    animation: {
                        duration: 2000,
                        easing: 'out'
                    },
                    majorTicks: ['', '', '', '', '$25000', '', '', '', '', '$50000'],  // Custom tick marks
                    fontSize: 14  // Reduce font size for better readability
                };
    
                var attendingOptions = {
                    width: 250, height: 150,
                    redFrom: 0, redTo: 49,
                    yellowFrom: 50, yellowTo: 124,
                    greenFrom: 125, greenTo: 250,
                    minorTicks: 5,
                    max: 250,  // Adjusted max range
                    animation: {
                        duration: 2000,
                        easing: 'out'
                    },
                    fontSize: 14  // Reduce font size for better readability
                };
    
                var outstandingChart = new google.visualization.Gauge(document.getElementById('outstanding_chart'));
                var attendingChart = new google.visualization.Gauge(document.getElementById('attending_chart'));
    
                outstandingChart.draw(outstandingData, options);
                attendingChart.draw(attendingData, attendingOptions);
    
                // Animate once to final values
                setTimeout(function() {
                    var outstandingValue = ''' + str(TotalOutstanding) + ''';  // Example final outstanding value
                    var attendingValue = ''' + str(TotalMembers) + ''';  // Example final attending count
    
                    outstandingData.setValue(0, 1, outstandingValue);
                    attendingData.setValue(0, 1, attendingValue);
    
                    outstandingChart.draw(outstandingData, options);
                    attendingChart.draw(attendingData, attendingOptions);
    
                    // Update number display manually with currency format
                    document.getElementById('outstanding_value').innerText = `$${outstandingValue.toLocaleString()}`;
                }, 500);  // Small delay before animation
            }
        </script>
        <div class="container">
        <div class="maincolumna">
            <p>
            <h3>Trips</h3>
            <table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 800px; width: 100%;">
                <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                    <td style="padding: 2px;"><h4>Reg.</h4></td>
                    <td style="padding: 2px;"><h4>Involvement</h4></td>
                    <td style="padding: 2px;"><h4>Going</h4></td>
                    <td style="padding: 2px;"><h4>Outstanding</h4></td>
                </tr>
    '''
    
    rsqlActiveMissions = q.QuerySql(sqlActiveMissions)
    
    for am in rsqlActiveMissions:
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{0}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;"><a href="?OrgView={1}">{2}</a></td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{3}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{4}</td>
            </tr>
            '''.format(str(am.RegistrationStatus),str(am.OrganizationId),am.OrganizationName,str(am.MemberCount),str(am.Outstanding))
        
    
    print '''</table><hr>
        <h3>Individual Due</h3>
        <table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 800px; width: 100%;">
            <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                <td style="padding: 2px;"><h4>Involvement</h4></td>
                <td style="padding: 2px;"><h4>Person</h4></td>
                <td style="padding: 2px;"><h4>Paid</h4></td>
                <td style="padding: 2px;"><h4>Outstanding</h4></td>
            </tr>
        '''
    
    rsqlOutstandingPayments = q.QuerySql(sqlOutstandingPayments)
    
    for op in rsqlOutstandingPayments:
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{0}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{1}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{2}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">{3}</td>
            </tr>
            '''.format(op.OrganizationName,op.Name2,str(op.Paid),str(op.Outstanding))
            
    print '''</table></div>
        <div class="rightcolumna"><p>
        </p>
            <h3>Stats</h3>
            <table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 200px; width: 100%;">
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
        <table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 200px; width: 100%;">
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
        <table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 200px; width: 100%;">
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
        <table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 200px; width: 100%;">
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
            -- Aggregate transaction data to avoid duplication
            FORMAT(ISNULL(ts_data.TotalPaid, 0), 'C', 'en-US') AS [Paid],
            FORMAT(ISNULL(ts_data.TotalCoupons, 0), 'C', 'en-US') AS [Coupons],
            FORMAT(ISNULL(ts_data.TotalOutstanding, 0), 'C', 'en-US') AS [Outstanding]
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
        -- Pre-aggregate transaction data per person to avoid duplication
        OUTER APPLY (
            SELECT 
                ts.PeopleId,
                SUM(ts.TotPaid) AS TotalPaid,
                SUM(ts.TotCoupon) AS TotalCoupons,
                SUM(ts.IndDue) AS TotalOutstanding
            FROM TransactionSummary ts
            WHERE ts.OrganizationId = o.OrganizationId
            AND ts.PeopleId = p.PeopleId  -- Ensuring transactions match the person
            GROUP BY ts.PeopleId
        ) ts_data
        WHERE o.OrganizationId = {0} -- Ensure you're filtering for a specific org
        ORDER BY mt.Description,p.Name2;

    
    '''

    print '''<button onclick="history.back()">Go Back</button>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
    '''
    
    rsqlInvolvementInformation = q.QuerySql(sqlInvolvementInformation.format(str(model.Data.OrgView)))
    for ii in rsqlInvolvementInformation:
        print '''<img src="{0}" alt="Mission Image Missing" style="width: 100%; height: auto;"><hr>
            <h2><a href="/Org/{6}" target="_blank"><i class="fa-solid fa-people-roof"></i></a>&nbsp{1}</h2>
            <strong>Status:</strong> {2}<br>
            <strong>Total Outstanding:</strong> {3}<br>
            <strong>Registration Start:</strong> {4}<br>
            <strong>Registration End:</strong> {5}<br>
            '''.format(ii.ImageUrl,ii.OrganizationName,ii.RegistrationStatus,str(ii.Outstanding),str(ii.RegStart),str(ii.RegEnd),str(ii.OrganizationId))
    
    print '''
            <h3>Meetings</h3>
            <hr>
            <h3>Team</h3>
            <table id="trnoborder" style="border: 0px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1200px; width: 100%;">
                <tr id="trnoborder" style="border: 0px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1200px; width: 100%;">
                    <td style="padding: 2px;"><h4>Name (Age)</h4></td>
                    <td style="padding: 2px;"><h4>Info</h4></td>
                    <td style="padding: 2px;"><h4>CheckList</h4></td>
                    <td style="padding: 2px;"><h4>Outstanding</h4></td>
                </tr>
            '''
    orgMemType = ''    
    rsqlInvolvementMembers = q.QuerySql(sqlInvolvementMembers.format(str(model.Data.OrgView)))
    for im in rsqlInvolvementMembers:
        formatted_homephone = '☎️ ' + model.FmtPhone(im.HomePhone) + '<br>' if im.HomePhone and im.HomePhone != im.CellPhone else ''
        
        if orgMemType == '' or orgMemType != im.OrgMemType:
            orgMemType = im.OrgMemType
            print '''<tr style="background-color: #004085; color: white; font-size: 18px; font-weight: bold; text-align: center;">
                        <td colspan="5" style="padding: 8px;">
                            {0}
                        </td>
                    </tr>
                 
                    '''.format(im.OrgMemType)
        print '''
            <tr>
                <td style="border-top: 1px dashed lightblue; padding: 2px;"><a href="/Person2/{11}" target="_blank"><i class="fa fa-info-circle"></i></a>&nbsp{1} ({2})</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">📱 {5}</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">% Checks:</td>
                <td style="border-top: 1px dashed lightblue; padding: 2px;">Outstanding: {7}</td>
            </tr>
            <tr>
                <td><img src="{12}" alt="Need Picture"></td> <!-- Empty cell for alignment -->
                <td colspan="1" style="padding: 2px; font-size: 14px; color: gray; vertical-align: top;">
                    {6}
                    {10}<br>
                    Baptism: {3}<br>
                    Church: {4}<hr style="margin: 0; padding: 0; border: none; border-top: 1px solid #999;">
                    Emergency Contact: {13}<br>
                    Emergency Phone: {14}
                </td>

                <td colspan="1" style="padding: 2px; font-size: 14px; color: gray; vertical-align: top;">
                    Future Check 1:<br>
                    Future Check 2:<br>
                </td> 
                <td colspan="1" style="padding: 2px; font-size: 14px; color: gray; vertical-align: top;">
                    Coupons: {8}<br>
                    Paid: {9}
                </td>
            </tr>

            '''.format(im.OrgMemType 
                        ,im.Name2 
                        ,im.Age
                        ,im.BaptismStatus
                        ,im.MemberStatus
                        ,model.FmtPhone(im.CellPhone)
                        ,formatted_homephone
                        ,im.Outstanding
                        ,im.Paid,im.Coupons
                        ,im.EmailAddress
                        ,im.PeopleId
                        ,im.Picture
                        ,im.emcontact
                        ,im.emphone)
            
            #                <!-- Contact Row (Indented slightly for readability) -->
            #    <tr>
            #        <td></td> <!-- Empty cell for alignment -->
            #        <td colspan="1" style="padding: 2px; font-size: 14px; color: gray;">
            #            📱 {5} | ☎️ {6} 
            #        </td>
            #    </tr>
            #| 📧 {10}
            
    print '''</table>'''
