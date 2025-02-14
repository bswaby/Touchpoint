#####################################################################
####REPORT INFORMATION
#####################################################################
#Contribution totals is a tool to aide financials in verifying their books.  
#To use this tool just create a Python script by going to Admin ~ 
#Advanced ~ Special Content and then click on the python tab

#####################################################################
####USER CONFIG FIELDS
#####################################################################
#User config is done within the script itself.  Just copy this script to 
#a new Python file and run.  The script will prompt if a config file has
#not been created

#####################################################################
####START OF CODE.  No configuration should be needed beyond this point
#####################################################################

#global

######################
##Config File Section
######################
from datetime import datetime
import re


#remove negative values
#def checkNeg(value):
#    #if value == -0.0:
#    #    return "-"
#    #else:
#    #    return "${:,.2f}".format(value)
#    return str(value)

def checkNeg(value):
    # Handle -0.0 explicitly
    if value == 0.0 and str(value).startswith('-'):
        return "-"
    
    # Handle 0.0 and return "-"
    if value == 0.0:
        return "-"
    
    # Check if the value is negative
    if value < 0:
        # Format negative numbers with parentheses
        return "(${:,.2f})".format(-value)  # Use -value to make it positive for formatting
    else:
        # Format positive numbers normally
        return "${:,.2f}".format(value)


def add_space_before_capital(text):
    return re.sub(r'(?<!^)(?=[A-Z])', ' ', text)

#config file name
ConfigFile = 'ConfigFinanceSummaryReport'

#Pull config text file
data = model.DynamicDataFromJson(model.TextContent(ConfigFile))    

#If text file doesn't exit, present message nad option to set with default values
json = ''
if not data:
    json = '''
            {
                "Script": {
                    "HeaderName": "Add title here",
                    "LastUpdatedBy": "",
                    "UpdatedLast": ""
                },
                "Email": {
                    "FromName": "John Smith",
                    "FromAddress": "jsmith@noreply.com",
                    "Subject": "Add Email Title Here"
                },
                "Financial": {
                    "ExcludedFundIds": "1307",
                    "GeneralFundIds": "1",
                    "CashBundleIds": "2",
                    "CheckBundleIds": "1,33",
                    "DepositTotalIds: "6.7",
                    "NonContributionTypeIds": "99",
                    "BundleReportName": "BundleReport3"
                },
                "Preferences": {
                    "ShowDepositTotals": "yes",
                    "ShowFund": "yes",
                    "ShowBundleSummary": "yes",
                    "ShowBundleDetails": "yes"
                }
            }
        '''
    data = model.DynamicDataFromJson(json)

#############Determine if values are from url or json.
#Script
HeaderName = str(model.Data.HeaderName) if model.Data.HeaderName else str(data.Script.HeaderName)
LastUpdatedBy = model.UserName
UpdatedLast = datetime.now()
#Email
FromName = str(model.Data.FromName) if model.Data.FromName else str(data.Email.FromName)
FromAddress = str(model.Data.FromAddress) if model.Data.FromAddress else str(data.Email.FromAddress)
Subject = str(model.Data.Subject) if model.Data.Subject else str(data.Email.Subject)
#Financial
ExcludedFundIds = str(model.Data.ExcludedFundIds) if model.Data.ExcludedFundIds else str(data.Financial.ExcludedFundIds)
GeneralFundIds = str(model.Data.GeneralFundIds) if model.Data.GeneralFundIds else str(data.Financial.GeneralFundIds)
CashBundleIds = str(model.Data.CashBundleIds) if model.Data.CashBundleIds else str(data.Financial.CashBundleIds)
CheckBundleIds = str(model.Data.CheckBundleIds) if model.Data.CheckBundleIds else str(data.Financial.CheckBundleIds)
DepositTotalIds = str(model.Data.DepositTotalIds) if model.Data.DepositTotalIds else str(data.Financial.DepositTotalIds)
NonContributionTypeIds = str(model.Data.NonContributionTypeIds) if model.Data.NonContributionTypeIds else str(data.Financial.NonContributionTypeIds)
BundleReportName = str(model.Data.BundleReportName) if model.Data.BundleReportName else str(data.Financial.BundleReportName)
#FormPreferences
ShowDepositTotals = str(model.Data.ShowDepositTotals) if model.Data.ShowDepositTotals else str(data.Preferences.ShowDepositTotals)
ShowFund = str(model.Data.ShowFund) if model.Data.ShowFund else str(data.Preferences.ShowFund)
ShowBundleSummary = str(model.Data.ShowBundleSummary) if model.Data.ShowBundleSummary else str(data.Preferences.ShowBundleSummary)
ShowBundleDetails = str(model.Data.ShowBundleDetails) if model.Data.ShowBundleDetails else str(data.Preferences.ShowBundleDetails)
ShowDepositTotalsChecked = 'checked' if ShowDepositTotals == "yes" else  '' 
ShowFundChecked = 'checked' if ShowFund == "yes" else ''
ShowBundleSummaryChecked = 'checked' if ShowBundleSummary == "yes" else ''
ShowBundleDetailsChecked = 'checked' if ShowBundleDetails == "yes" else ''


#print 'payment:' + ShowPaymentType + 'fund:' + ShowFund + 'summary:' + ShowBundleSummary
#print '<br>payment:' + ShowPaymentTypeChecked + 'fund:' + ShowFundChecked + 'summary:' + ShowBundleSummaryChecked


print '''  

        <style>
            /* Popup container */
            .popup {
                display: none;
                position: fixed;
                z-index: 1;
                left: 0;
                top: 0;
                width: 100%;
                height: 100%;
                overflow: auto;
                background-color: rgba(0,0,0,0.4);
            }
    
            /* Popup content */
            .popup-content {
                background-color: #fefefe;
                margin: 15% auto;
                padding: 20px;
                border: 1px solid #888;
                width: 80%;
            }
    
            /* Show the popup when the URL has #popup */
            .popup:target {
                display: block;
            }
        
            #trnoborder {
              border: 1px solid #dddddd;
              text-align: left;
              padding: 10px;
              font-family: arial, sans-serif;
              border-collapse: collapse;
              max-width: 1000px; /* Maximum width of the table */
              width: 100%; /* Table adjusts to fit container */
            }
            
            #trFundnoborder {
              border: 1px solid #dddddd;
              text-align: left;
              padding: 10px;
              font-family: arial, sans-serif;
              border-collapse: collapse;
              max-width: 600px; /* Maximum width of the table */
              width: 100%; /* Table adjusts to fit container */
            }
            
            tr:nth-child(even) {e
            
            }
            td {
                padding: 0 4px 0 4px; /* Top, Right, Bottom, Left */
            }
            .left-border {
                border-left: 1px solid black;
                background-color: lightblue;
            }
            .topleft-border {
                border-left: 1px solid black;
                border-top: 1px solid black;
                background-color: lightblue;
            }
            .topright-border {
                border-right: 1px solid black;
                border-top: 1px solid black;
                background-color: lightblue;
            }
            .leftright-border {
                border-left: 1px solid black;
                border-right: 1px solid black;
                background-color: lightblue;
            }
            .leftrighttop-border {
                border-left: 1px solid black;
                border-right: 1px solid black;
                border-top: 1px solid black;
                background-color: lightblue;
            }
            .top-borderdash {
                border-bottom: 1px dashed lightblue;
            }
            .top-border {
                border-top: 1px solid black;
            }
            .total-cell {
                background-color: lightblue;
                border-right: 1px solid black;
            }
            /* Style to display checkboxes inline */
            .checkbox-container {
                display: inline-block;
            }
            input[type="checkbox"] {
                display: inline;
            }
            body {
                -webkit-print-color-adjust: exact; /*Chrome,Safari,Edge*/
                color-adjust: exact;              /* firefox*/
             }
            @media print {
                td.total-column {
                    background-color: lightblue !important;
                    background-image: linear-gradient(lightblue, lightblue) !important;
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
            }
            @media print {
                td.subtotal-column {
                    background-color: #f2f2f2 !important;
                    background-image: linear-gradient(#f2f2f2, #f2f2f2) !important;
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
            }
            @media print {
                td.gray-text {
                    color: black !important; /* Keep it black but make it appear gray */
                    opacity: 0.6 !important; /* Adjust this for lighter or darker gray */
                }
            }

            @media print {
                .no-print {
                    display: none !important;
                }
            }
        </style>'''

if model.Data.EditParams == 'y':

    #Set json file to write back
    Newjson = '''
    {
        "Script": {
            "HeaderName": "''' + HeaderName + '''",
            "LastUpdatedBy": "''' + LastUpdatedBy + '''",
            "UpdatedLast": "''' + str(UpdatedLast) + '''"
        },
        "Email": {
            "FromName": "''' + FromName + '''",
            "FromAddress": "''' + FromAddress + '''",
            "Subject": "''' + Subject + '''"
        },
        "Financial": {
            "ExcludedFundIds": "''' + ExcludedFundIds + '''",
            "GeneralFundIds": "''' + GeneralFundIds + '''",
            "CashBundleIds": "''' + CashBundleIds + '''",
            "CheckBundleIds": "''' + CheckBundleIds + '''",
            "DepositTotalIds": "''' + DepositTotalIds + '''",
            "NonContributionTypeIds": "''' + NonContributionTypeIds + '''",
            "BundleReportName": "''' + BundleReportName + '''"
        },
        "Preferences": {
            "ShowDepositTotals": "''' + ShowDepositTotals + '''",
            "ShowFund": "''' + ShowFund + '''",
            "ShowBundleSummary": "''' + ShowBundleSummary + '''",
            "ShowBundleDetails": "''' + ShowBundleDetails + '''"
        }
    }
    '''

    #write json file back
    model.WriteContentText(ConfigFile, Newjson)
    print '''
        <!-- Go Back Button -->
        <button onclick="goBack()">Go Back</button>
        
        <script>
        // Function to remove query parameters and anchors from a URL
        function cleanUrl(url) {
          var cleanUrl = url.split('?')[0]; // Remove query parameters (if any)
          cleanUrl = cleanUrl.split('#')[0]; // Remove anchor (if any)
          return cleanUrl;
        }
        
        // Function for going back to the previous page with a cleaned URL
        function goBack() {
          var referrer = document.referrer; // Get the previous page's URL
          if (referrer) {
            // Clean the referrer URL (remove query params and anchors)
            var cleanedReferrer = cleanUrl(referrer);
            // Go to the cleaned URL
            window.location.href = cleanedReferrer;
          } else {
            // If there is no referrer, just go back using history.back()
            window.history.back();
          }
        }
        </script>
        <br>
        <h3>Parameters Updated</h3>'''
else:
    if json != '':
        print '''<h3>Config file does not exist.  You must set it first before script will run.</h2>'''
    #removed params
    #<input type="hidden" value="no" name='ShowBundleSummary'>
    #<label for="ExcludedFundIds">Excluded Funds:</label>
    #<input type="text" id="ExcludedFundIds" name="ExcludedFundIds" value="{4}"><br>
    print '''
        <!-- Link to open the popup -->
        <a href="#popup" class="no-print">Edit Parameters</a>
        
        <!-- The popup container -->
        <div id="popup" class="popup">
            <!-- Popup content -->
            <div class="popup-content">
                <a href="#">Close</a>
    
                <form id="editParams" action="" method="get">
                    <input type="hidden" name="EditParams" value="y">
                    <h4>Form Preferences</h4>
                    <label for="email">Form Name:</label>
                    <input type="text" id="HeaderName" name="HeaderName" value="{0}"><br>
                    <label>Show/Hide:</label>
                    <div class="checkbox-container">
                        <input type="checkbox" id="ShowBundleSummary" name="ShowBundleSummary" value="yes" {9}>
                        <label for="ShowBundleSummary">Bundle Summary</label>
                    </div>
            
                    <div class="checkbox-container">
                        <input type="checkbox" id="ShowFund" name="ShowFund" value="yes" {10}>
                        <label for="ShowFund">Fund</label>
                    </div>
            
                    <div class="checkbox-container">
                        <input type="checkbox" id="ShowDepositTotals" name="ShowDepositTotals" value="yes" {11}>
                        <label for="ShowDepositTotals">Deposit Totals</label>
                    </div>
                    <div class="checkbox-container">
                        <input type="checkbox" id="ShowBundleDetails" name="ShowBundleDetails" value="yes" {12}>
                        <label for="ShowBundleDetails">Bundle Details</label>
                    </div>

                    <hr>
                    <h4>Email Parameters</h4>
                    <label for="FromName">From Name:</label>
                    <input type="text" id="FromName" name="FromName" value="{1}"><br>
                    <label for="FromAddress">From Address:</label>
                    <input type="text" id="FromAddress" name="FromAddress" value="{2}"><br>
                    <label for="Subject">Email Subject:</label>
                    <input type="text" id="Subject" name="Subject" value="{3}"><br>
                    <hr>
                    <h4>Financial Parameters</h4>
                    <i>seperate multiple Id's by a comma.  example: 1,11,12</i></br>
                    <label for="GeneralFundIds">General Fund:</label>
                    <input type="text" id="GeneralFundIds" name="GeneralFundIds" value="{5}"><br>
                    <label for="DepositTotalIds">Contribution Source Ids:</label>
                    <input type="text" id="DepositTotalIds" name="DepositTotalIds" value="{16}"><br>
                    <label for="NonContributionTypeIds">NonContribution Type Id:</label>
                    <input type="text" id="NonContributionTypeIds" name="NonContributionTypeIds" value="{15}"><br>
                    <label for="BundleReportName">Bundle Report Name:</label>
                    <input type="text" id="Subject" name="BundleReportName" value="{6}"><br>
                    <i>- Bundle report can be found under Custom Batch Report in Admin ~ Settings ~ Finance ~ Batches<br>
                     - NonContribution Type ID's and BundleID's can be found under Admin ~ Advanced ~ Lookup Codes </i>
                    </br>
                    <input type="submit" value="Submit">
                    <br><br><i>Last updated by {7} on {8}</i>
                </form>
            </div>
        </div>
    '''.format(HeaderName,
                FromName,
                FromAddress,
                Subject,
                ExcludedFundIds,
                GeneralFundIds,
                BundleReportName,
                LastUpdatedBy,
                UpdatedLast,
                ShowBundleSummaryChecked,
                ShowFundChecked,
                ShowDepositTotalsChecked,
                ShowBundleDetailsChecked,
                CashBundleIds,
                CheckBundleIds,
                NonContributionTypeIds,
                DepositTotalIds)
    print '''  <script>
          document.getElementById("editParams").onsubmit = function() {
            var checkbox = document.getElementById("ShowBundleSummary");
            if (!checkbox.checked) {
              // If checkbox is unchecked, set its value to 'no' before submitting
              var input = document.createElement("input");
              input.type = "hidden";
              input.name = "ShowBundleSummary";
              input.value = "no";
              this.appendChild(input);  // Append the hidden input to the form
            }
            var checkbox = document.getElementById("ShowFund");
            if (!checkbox.checked) {
              // If checkbox is unchecked, set its value to 'no' before submitting
              var input = document.createElement("input");
              input.type = "hidden";
              input.name = "ShowFund";
              input.value = "no";
              this.appendChild(input);  // Append the hidden input to the form
            }
            var checkbox = document.getElementById("ShowDepositTotals");
            if (!checkbox.checked) {
              // If checkbox is unchecked, set its value to 'no' before submitting
              var input = document.createElement("input");
              input.type = "hidden";
              input.name = "ShowDepositTotals";
              input.value = "no";
              this.appendChild(input);  // Append the hidden input to the form
            }
            var checkbox = document.getElementById("ShowBundleDetails");
            if (!checkbox.checked) {
              // If checkbox is unchecked, set its value to 'no' before submitting
              var input = document.createElement("input");
              input.type = "hidden";
              input.name = "ShowBundleDetails";
              input.value = "no";
              this.appendChild(input);  // Append the hidden input to the form
            }
          };
        </script>'''
        
    model.Header = HeaderName#"Weekly Contribution Totals" #Set to the name you want to call the page

    # Get URL Variables
    import re
    import locale
    from types import NoneType
    import datetime


    #Get URL Variables
    sDate = model.Data.sDate
    eDate = model.Data.eDate
    #DateSearch = model.Data.DateSearch
    sendReport = model.Data.sendReport
    ShowDesignatedDetails = model.Data.ShowDesignatedDetails
    HideBundle = model.Data.HideBundle
    FundSort = model.Data.FundSort

    
    #set form parameters 
    if model.Data.DateSearch == 'ContributionDate':
        ContributionDateChecked = 'checked'
        DepositDateChecked = ''
        DateSearch = 'cb.ContributionDate '
    elif model.Data.DateSearch == 'DepositDate':
        ContributionDateChecked = ''
        DepositDateChecked = 'checked'
        DateSearch = 'bl.DepositDate '
    else:
        ContributionDateChecked = ''
        DepositDateChecked = 'checked'
        DateSearch = 'bl.DepositDate '

       
    if sDate is not None:
        optionsDate = ' value="' + sDate + '"'
        searchsDate = sDate
    else:
        searchsDate = '2024-10-29'
    
    if eDate is not None:
        optioneDate = ' value="' + eDate + '"'
        searcheDate = eDate
    else:
        searcheDate = '2024-11-04'

    if HideBundle == 'yes':
        optionHideBundle = 'checked'
    else:
        optionHideBundle = ''
    
    if ShowDesignatedDetails == 'yes':
        optionShowDesignatedDetails = 'checked'
    else:
        optionShowDesignatedDetails = ''

    # Use dictionary for FundSort selection (Python 2 compatible)
    fund_options = {
        'Fund': ('checked', ''),
        'FundSetSort': ('', 'checked')
    }
    optionFundSort, optionFundSetSort = fund_options.get(FundSort, ('', 'checked'))
        
    
    sql = '''
    
        WITH bundlereport AS (
        SELECT 
            bl.[status] AS BundleStatus,
            bl.HeaderType,
            cs.Description AS ContributionType,
            CASE 
                WHEN cb.FundId IN ({4}) THEN SUM(cb.ContributionAmount) 
                ELSE 0 
            END AS General,
            CASE 
                WHEN ct.id NOT IN ({5}) AND cb.FundId NOT IN ({4}) THEN SUM(cb.ContributionAmount) 
                ELSE 0 
            END AS Designated,
            CASE 
                WHEN ct.id IN ({5}) THEN SUM(cb.ContributionAmount) 
                ELSE 0 
            END AS NonContribution
        FROM [ContributionsBasic] cb
        LEFT JOIN BundleDetail bd ON cb.ContributionId = bd.ContributionId
        LEFT JOIN BundleList bl ON bl.BundleHeaderId = bd.BundleHeaderId
        LEFT JOIN BundleHeader bh ON bh.BundleHeaderId = bl.BundleHeaderId
        LEFT JOIN lookup.ContributionType ct ON ct.Id = cb.ContributionTypeId
        LEFT JOIN ContributionFund cf ON cf.FundId = cb.FundId
        LEFT JOIN lookup.ContributionSources cs ON cs.Id = bh.SourceId
        --WHERE bl.DepositDate BETWEEN '2025-01-01' AND '2025-01-12'
        Where {2} BETWEEN '{0}' AND '{1}'
        GROUP BY 
            bl.HeaderType, 
            cb.BundleReferenceId, 
            ct.[Description], 
            ct.Id, 
            cb.FundId, 
            bd.BundleHeaderId, 
            bl.[status], 
            cs.Description
    )
    SELECT 
        br.BundleStatus,
        br.HeaderType,
        COUNT(DISTINCT br.ContributionType) AS ContributionTypeCount,
        (
            SELECT STRING_AGG(ContributionType, ',')
            FROM (
                SELECT DISTINCT ContributionType
                FROM bundlereport sub
                WHERE sub.BundleStatus = br.BundleStatus
                  AND sub.HeaderType = br.HeaderType
            ) DistinctTypes
        ) AS CommaSeparatedList,
        SUM(br.General) AS General,
        SUM(br.Designated) AS Designated,
        SUM(br.NonContribution) AS NonContribution,
        SUM(br.General) + SUM(br.Designated) AS TotalContribution,
        SUM(br.General) + SUM(br.Designated) + SUM(br.NonContribution) AS Total
    FROM bundlereport br
    GROUP BY br.HeaderType, br.BundleStatus
    ORDER BY br.HeaderType, br.BundleStatus;

    
    '''
    sqlbundledetail = '''
    WITH bundlereport AS (    
        SELECT 
        bl.[status] AS BundleStatus,
        bl.HeaderType,
        cs.Description AS ContributionType,
        CASE 
            WHEN cb.FundId IN ({4}) THEN SUM(cb.ContributionAmount) 
            ELSE 0 
        END AS General,
        CASE 
            WHEN ct.id NOT IN ({5}) AND cb.FundId NOT IN ({4}) THEN SUM(cb.ContributionAmount) 
            ELSE 0 
        END AS Designated,
        CASE 
            WHEN ct.id IN ({5}) THEN SUM(cb.ContributionAmount) 
            ELSE 0 
        END AS NonContribution
    FROM [ContributionsBasic] cb
    LEFT JOIN BundleDetail bd ON cb.ContributionId = bd.ContributionId
    LEFT JOIN BundleList bl ON bl.BundleHeaderId = bd.BundleHeaderId
    LEFT JOIN BundleHeader bh ON bh.BundleHeaderId = bl.BundleHeaderId
    LEFT JOIN lookup.ContributionType ct ON ct.Id = cb.ContributionTypeId
    LEFT JOIN ContributionFund cf ON cf.FundId = cb.FundId
    LEFT JOIN lookup.ContributionSources cs ON cs.Id = bh.SourceId
    WHERE {2} BETWEEN '{0}' AND '{1}'
	AND bl.[status] = '{6}'
	AND bl.HeaderType = '{7}'
    AND cs.Description = '{8}'
	GROUP BY 
        bl.HeaderType, 
        cb.BundleReferenceId, 
        ct.[Description], 
        ct.Id, 
        cb.FundId, 
        bd.BundleHeaderId, 
        bl.[status], 
        cs.Description
    )
    
    SELECT 
        br.BundleStatus,
        br.HeaderType,
        COUNT(DISTINCT br.ContributionType) AS ContributionTypeCount,
        (
            SELECT STRING_AGG(ContributionType, ',')
            FROM (
                SELECT DISTINCT ContributionType
                FROM bundlereport sub
                WHERE sub.BundleStatus = br.BundleStatus
                  AND sub.HeaderType = br.HeaderType
            ) DistinctTypes
        ) AS CommaSeparatedList,
        SUM(br.General) AS General,
        SUM(br.Designated) AS Designated,
        SUM(br.NonContribution) AS NonContribution,
        SUM(br.General) + SUM(br.Designated) AS TotalContribution,
        SUM(br.General) + SUM(br.Designated) + SUM(br.NonContribution) AS Total
    FROM bundlereport br
    GROUP BY br.HeaderType, br.BundleStatus
    ORDER BY br.HeaderType, br.BundleStatus;            
    '''

    
    sqlBundle = '''
        with bundlereport as (
        Select 
            bl.HeaderType
            ,case when cb.FundId in ({4}) THEN sum(cb.ContributionAmount) ELSE 0 END AS General
            ,case when ct.id NOT IN ({5}) and cb.fundId not in ({4}) THEN sum(cb.ContributionAmount) ELSE 0 End as Designated
            ,case when ct.id IN ({5}) THEN sum(cb.ContributionAmount) ELSE 0 End as NonContribution
        FROM [ContributionsBasic] cb
            LEFT JOIN 
                BundleDetail bd ON cb.ContributionId = bd.ContributionId
            LEFT JOIN 
                BundleList bl ON bl.BundleHeaderId = bd.BundleHeaderId
            LEFT JOIN 
                lookup.ContributionType ct ON ct.Id = cb.ContributionTypeId
            LEFT JOIN 
                ContributionFund cf ON cf.FundId = cb.FundId
            --Where ContributionDate BETWEEN '{0}' AND '{1}'
        Where {2} BETWEEN '{0}' AND '{1}'
        Group By 
            bl.HeaderType
            ,cb.BundleReferenceId
            ,ct.[Description]
            ,ct.Id
            ,cb.FundId
            ,bd.BundleHeaderId
        )
        
        Select HeaderType
        ,sum(General) as General
        ,sum(Designated) as Designated
        ,sum(NonContribution) as NonContribution
        ,sum(General) + sum(Designated)  AS TotalContribution
        ,sum(General) + sum(Designated) + sum(NonContribution) AS Total
        From bundlereport
        Group By HeaderType
    '''
    
    sqlDetails = '''
        with bundlereport as (
        Select 
             bl.HeaderType
            ,bd.BundleHeaderId
            ,bst.[Description] AS BundleStatus
            ,FORMAT(bl.DepositDate, 'yyyy-MM-dd') as DepositDate
            ,concat(cb.BundleReferenceId,' (',bd.BundleHeaderId,')') AS ReferenceId
            ,case when cb.FundId IN ({4}) THEN sum(cb.ContributionAmount) ELSE 0 END AS General
            ,case when ct.id NOT IN ({6}) and cb.fundId not in ({4})  THEN sum(cb.ContributionAmount) ELSE 0 End as Designated
            ,case when ct.id IN ({6}) THEN sum(cb.ContributionAmount) ELSE 0 End as NonContribution
        FROM [ContributionsBasic] cb
            LEFT JOIN BundleDetail bd ON cb.ContributionId = bd.ContributionId
            LEFT JOIN BundleList bl ON bl.BundleHeaderId = bd.BundleHeaderId
            LEFT JOIN lookup.ContributionType ct ON ct.Id = cb.ContributionTypeId
            LEFT JOIN ContributionFund cf ON cf.FundId = cb.FundId
            LEFT JOIN lookup.BundleStatusTypes bst ON bst.Id = bl.BundleStatusId
        --Where ContributionDate BETWEEN '{0}' AND '{1}' 
        Where {5} BETWEEN '{0}' AND '{1}' 
        and bl.HeaderType = '{2}'
        Group By bl.HeaderType,cb.BundleReferenceId,ct.[Description],ct.Id,cb.FundId,bd.BundleHeaderId,bst.[Description],bl.DepositDate
        )
        
        Select 
             ReferenceId
            ,BundleHeaderId
            ,HeaderType
            ,BundleStatus
            ,DepositDate
            ,sum(General) as General
            ,sum(Designated) as Designated
            ,sum(NonContribution) as NonContribution
            ,sum(General) + sum(Designated)  AS TotalContribution
            ,sum(General) + sum(Designated) + sum(NonContribution) AS Total
        From bundlereport
        Group By HeaderType,ReferenceId,BundleHeaderId,BundleStatus,DepositDate
    '''
    
    sqlFundDetails = '''
        with bundlereport as (
        Select 
        bl.HeaderType
        ,concat(cf.FundName,' (',cb.FundId,')') AS Fund
        ,bd.BundleHeaderId
        ,bst.[Description] AS BundleStatus
        ,concat(cb.BundleReferenceId,' (',bd.BundleHeaderId,')') AS ReferenceId
        ,case when cb.FundId IN ({3}) THEN sum(cb.ContributionAmount) ELSE 0 END AS General
        ,case when ct.id NOT IN ({4}) and cb.fundId not in ({3}) THEN sum(cb.ContributionAmount) ELSE 0 End as Designated
        ,case when ct.id IN ({4}) THEN sum(cb.ContributionAmount) ELSE 0 End as NonContribution
        FROM [ContributionsBasic] cb
        LEFT JOIN BundleDetail bd ON cb.ContributionId = bd.ContributionId
        LEFT JOIN BundleList bl ON bl.BundleHeaderId = bd.BundleHeaderId
        LEFT JOIN lookup.ContributionType ct ON ct.Id = cb.ContributionTypeId
        LEFT JOIN ContributionFund cf ON cf.FundId = cb.FundId
        LEFT JOIN lookup.BundleStatusTypes bst ON bst.Id = bl.BundleStatusId
        Where bd.BundleHeaderId = {0}
        Group By bl.HeaderType,cb.BundleReferenceId,ct.[Description],ct.Id,cb.FundId,bd.BundleHeaderId,cf.FundName,bst.[Description])
        
        Select 
        ReferenceId
        ,Fund
        ,BundleHeaderId
        ,HeaderType
        ,BundleStatus
        ,sum(General) as General
        ,sum(Designated) as Designated
        ,sum(NonContribution) as NonContribution
        ,sum(General) + sum(Designated)  AS TotalContribution
        ,sum(General) + sum(Designated) + sum(NonContribution) AS Total
        From bundlereport
        Group By HeaderType,ReferenceId,BundleHeaderId,BundleStatus,Fund
        Order By General Desc, Designated Desc, NonContribution Desc, HeaderType
    '''

    SqlFunds = '''
        WITH bundlereport AS (
            SELECT 
                bl.[Status] AS BundleStatus,
                CONCAT(cf.FundName, ' (', cb.FundId, ')') AS Fund,
        		case when cb.FundId in ({4})  THEN sum(cb.ContributionAmount) ELSE 0 END AS General,
        		case when ct.id NOT IN ({5}) and cb.fundId not in ({4})  THEN sum(cb.ContributionAmount) ELSE 0 End as Designated,
        		case when ct.id IN ({5}) THEN sum(cb.ContributionAmount) ELSE 0 End as NonContribution,
                fs.Description AS [FundSet]
            FROM [Contribution] cb
                LEFT JOIN BundleDetail bd ON cb.ContributionId = bd.ContributionId
                LEFT JOIN BundleList bl ON bl.BundleHeaderId = bd.BundleHeaderId
                LEFT JOIN lookup.ContributionType ct ON ct.Id = cb.ContributionTypeId
                LEFT JOIN ContributionFund cf ON cf.FundId = cb.FundId
                LEFT JOIN FundSetFunds fsf ON fsf.FundId = cb.FundId
                LEFT JOIN FundSets fs ON fs.FundSetId = fsf.FundSetId
        	Where {2} BETWEEN '{0}' AND '{1}'
            GROUP BY 
                bl.HeaderType, ct.[Description], ct.Id, cb.FundId, bd.BundleHeaderId, 
                cf.FundName, bl.[Status], fs.Description
        )
        
        SELECT 
            BundleStatus,
            Fund,
            SUM(General) AS General,
            SUM(Designated) AS Designated,
            SUM(NonContribution) AS NonContribution,
            FundSet,
            COUNT(*) OVER (PARTITION BY FundSet) AS FundSetCount  
        FROM bundlereport
        GROUP BY Fund, BundleStatus, FundSet
        ORDER BY 
            {6}
    '''
    
    SqlDepositTotalsOld = '''
        With DepositSum AS(
		Select bht.[Description] AS BundleType
		,CASE WHEN bht.Id IN ({2}) THEN sum(c.ContributionAmount) ELSE 0 END as Checks 
		,CASE WHEN bht.Id IN ({3}) THEN sum(c.ContributionAmount) ELSE 0 END as Cash 
        from BundleDetail bd
        left join Contribution c ON c.ContributionId = bd.ContributionId
        left join BundleHeader bh ON bd.BundleHeaderId = bh.BundleHeaderId
        left join lookup.BundleHeaderTypes bht ON bht.Id = bh.BundleHeaderTypeId
        where bh.DepositDate BETWEEN '{0}' AND '{1}'
        Group By bht.id, bht.[Description]
		)

		SELECT BundleType, sum(Checks) AS Checks, Sum(Cash) AS Cash From DepositSum Where Checks <> 0 Or Cash <> 0 Group By BundleType Order by BundleType
    '''
    
    SqlDepositTotals = '''
        select Sum(bh.BundleTotal) As [Total]
        	,cs.description as [Source]
        FROM BundleHeader bh
        	LEFT JOIN lookup.ContributionSources cs ON cs.Id = bh.SourceId
        where bh.DepositDate BETWEEN '{0}' AND '{1}'
         and bh.SourceId in ({2})
        Group By cs.Description
    '''

    template = ''

    # set form date fields and table headers for bundle summary
    ShowFundDetails = ''
    #if HideBundle == "yes":
    #    HideBundle = '''<input type="checkbox" id="HideBundle" name="HideBundle" value="yes" {0}>
    #        <label for="HideBundle">HideBundle</label>'''.format(optionHideBundle)
    if ShowBundleDetails == "yes":
        ShowFundDetails = '''<input class="no-print" type="checkbox" id="ShowDesignatedDetails" name="ShowDesignatedDetails" value="yes" {0}>
            <label for="ShowDesignatedDetails" class="no-print">Show Funds in Bundle Details</label>'''.format(optionShowDesignatedDetails)

    template += """ 
        <form action="" method="GET">
            <p class="no-print" style="margin: 0; padding: 0;">
                <label for="sDate">Start:</label>
                <input type="date" id="sDate" name="sDate" required {optionsDate}>
                <label for="eDate">End:</label>
                <input type="date" id="eDate" name="eDate" required {optioneDate}>
                <input type="submit" value="Filter">
            </p>
    
            <p class="no-print" style="margin: 0; padding: 0;">Search By:
                <label>
                    <input type="radio" name="DateSearch" value="DepositDate" {DepositDateChecked}>
                    Deposited Date
                </label>
                <label>
                    <input type="radio" name="DateSearch" value="ContributionDate" {ContributionDateChecked}>
                    Contributed Date
                </label>
            </p>
    
            <p style="margin: 0; padding: 0;" class="no-print">Sort Fund By:
                <label>
                    <input type="radio" class="no-print" name="FundSort" value="Fundset" {optionFundSetSort}>
                    Fund Set
                </label>
                <label>
                    <input type="radio" class="no-print" name="FundSort" value="Fund" {optionFundSort}>
                    Fund
                </label>
            </p>
    
            <p style="margin: 0; padding: 0;">
                <input class="no-print" type="checkbox" id="HideBundle" name="HideBundle" value="yes" {optionHideBundle}>
                <label for="HideBundle" class="no-print">Hide Bundle Details</label>
            </p>
            
            {ShowFundDetails}
        </form>
    
        <h2>{add_space_before} {sDate} - {eDate}</h2><br>
    """.format(
        optionsDate=optionsDate,
        optioneDate=optioneDate,
        ShowFundDetails=ShowFundDetails,
        DepositDateChecked=DepositDateChecked,
        ContributionDateChecked=ContributionDateChecked,
        optionHideBundle=optionHideBundle,
        sDate=sDate,
        eDate=eDate,
        add_space_before=add_space_before_capital(model.Data.DateSearch),
        optionFundSort=optionFundSort,
        optionFundSetSort=optionFundSetSort
    )
              
    #print '<h1>' + model.Data.DateSearch + '</h1>'
    
    if ShowBundleSummary == "yes":
        template += '''
            <h3>Bundle</h3>
            <table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                    <td><h4>Status</h4></td>
                    <td><h4>Type</h4></td>
                    <td><h4>Source</h4></td>
                    <td style="border-left: 1px dashed #dddddd;"><h4>General</h4></td>
                    <td><h4>Designated</h4></td>
                    <td class="subtotal-column" style="border-left: 1px dashed #dddddd; border-right: 1px dashed #dddddd;background-color: #f2f2f2;print-color-adjust: exact; -webkit-print-color-adjust: exact;"><h4>Contribution<br>Total</h4></td>
                    <td><h4>NC</h4></td>
                    <td class="total-column" style="background-color: lightblue;border-right: 1px solid #dddddd; border-left: 1px double #dddddd; print-color-adjust: exact; -webkit-print-color-adjust: exact;"><h4>Grand<br>Total</h4></td>
                </td>
        '''

    #bundle summary query
    sql = q.QuerySql(sql.format(searchsDate,searcheDate,DateSearch,ExcludedFundIds,GeneralFundIds,NonContributionTypeIds))
    
    # Set variables for bundle summary
    GrandTotalGenDes = 0.00
    GrandTotalGen = 0.00
    GrandTotalDes = 0.00
    GrandTotalNon = 0.00
    GrandTotal = 0.00
    
    # Loop through bundle summary
    for d in sql:
        # Row Totals
        TotalGenDes = d.General + d.Designated
    
        # Column Totals
        GrandTotalGenDes += TotalGenDes
        GrandTotalGen += d.General
        GrandTotalDes += d.Designated
        GrandTotalNon += d.NonContribution
        GrandTotal += d.General + d.Designated + d.NonContribution
    
        # Format Bundle Status
        BundleStatus = d.BundleStatus if d.BundleStatus == "Closed" else "<b>{}</b>".format(d.BundleStatus)
    
        # Contribution Type
        ContributionType = "" if d.ContributionTypeCount > 1 else (d.CommaSeparatedList or "")
    
        # Summary Details
        if ShowBundleSummary == "yes":
            template += """<tr>
                <td style="border-top: 1px dashed lightblue;">{BundleStatus}</td>
                <td style="border-top: 1px dashed lightblue;">{HeaderType}</td>
                <td style="border-top: 1px dashed lightblue;">{ContributionType}</td>
                <td style="border-top: 1px dashed lightblue; border-left: 1px dashed #dddddd;">{General}</td>
                <td style="border-top: 1px dashed lightblue;">{Designated}</td>
                <td class="subtotal-column" style="border-left: 1px dashed #dddddd; border-right: 1px dashed #dddddd; 
                    border-top: 1px dashed #dddddd; background-color: #f2f2f2; print-color-adjust:exact;">{TotalContribution}</td>
                <td style="border-top: 1px dashed lightblue;">{NonContribution}</td>
                <td class="total-column" style="border-left: 1px solid #dddddd; border-right: 1px solid #dddddd; 
                    background-color: lightblue; border-top: 1px dashed #dddddd; print-color-adjust:exact;">{Total}</td>
            </tr>""".format(
                BundleStatus=BundleStatus,
                HeaderType=d.HeaderType,
                ContributionType=ContributionType,
                General=checkNeg(d.General),
                Designated=checkNeg(d.Designated),
                NonContribution=checkNeg(d.NonContribution),
                TotalContribution=checkNeg(d.TotalContribution),
                Total=checkNeg(d.Total)
            )

            if d.ContributionTypeCount > 1:
                # Split the string into a list
                ContributionTypeitems = d.CommaSeparatedList.split(',')
                
                # Iterate through the list
                for item in ContributionTypeitems:
                    #print(item)
                    #print sqlbundledetail.format(searchsDate,searcheDate,DateSearch,ExcludedFundIds,GeneralFundIds,NonContributionTypeIds,BundleStatus,d.HeaderType,item)
                    rsqlbundledetail = q.QuerySql(sqlbundledetail.format(searchsDate,searcheDate,DateSearch,ExcludedFundIds,GeneralFundIds,NonContributionTypeIds,d.BundleStatus,d.HeaderType,item))
                    for rbd in rsqlbundledetail:
                        template += '''<tr>
                                            <td style="position: relative; background: linear-gradient(to right, transparent 49%, lightblue 49%, lightblue 51%, transparent 51%); background-size: 100% 2px;"></td>
                                            <td></td>
                                            <td class="gray-text" style="color: gray;"><i>{0}</i></td>
                                            <td class="gray-text" style="color: gray; border-left: 1px dashed #dddddd;"><i>&nbsp&nbsp&nbsp&nbsp{1}</i></td>
                                            <td class="gray-text" style="color: gray;"><i>&nbsp&nbsp&nbsp&nbsp{2}</i></td>
                                            <td class="subtotal-column" style="border-left: 1px dashed #dddddd; border-right: 1px dashed #dddddd; background-color: #f2f2f2;;print-color-adjust:exact;"></td>
                                            <td class="gray-text" style="border-left: 1px dotted #dddddd; border-right: 1px solid #dddddd; color: gray;"><i>&nbsp&nbsp&nbsp&nbsp{3}</i></td>
                                            <td class="total-column" style="background-color: lightblue;border-right: 1px solid #dddddd; border-left: 1px solid #dddddd; print-color-adjust:exact;"></td>
                                        </tr>
                                    '''.format(rbd.CommaSeparatedList,
                                               checkNeg(rbd.General),
                                               checkNeg(rbd.Designated),
                                               checkNeg(rbd.NonContribution))
        
    if ShowBundleSummary == "yes":    
        template += '''
                <tr>
                    <td><h4></h4></td>
                    <td><h4></h4></td>
                    <td><h4></h4></td>
                    <td style="border-top: 1px solid #000000; border-left: 1px dashed #dddddd;"><b>{0}</b></td>
                    <td style="border-top: 1px solid #000000;"><b>{1}</b></td>
                    <td style="border-left: 1px dashed #dddddd;border-right: 1px dashed #dddddd;border-top: 1px solid black;background-color: #f2f2f2;;print-color-adjust:exact;"><b>{3}</b></td>
                    <td style="border-left: 1px dotted #dddddd;border-right: 1px solid #dddddd;border-top: 1px solid black;"><b>{2}</b></td>
                    <td class="total-column" style="border-left: 1px solid #dddddd;border-right: 1px solid #dddddd; border-top: 1px solid black; background-color: lightblue;print-color-adjust:exact;"><b>{4}</b></td>
                </tr>
            </table>
            '''.format(checkNeg(GrandTotalGen),checkNeg(GrandTotalDes),checkNeg(GrandTotalNon),checkNeg(GrandTotalGenDes),checkNeg(GrandTotal))


    #Payment Summary    
    if ShowDepositTotals == "yes":    
        template +=  '''<h3>Deposit Totals</h3>
                
            <table id="trFundnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 600px;width: 100%;">
                <tr id="trFundnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 600px;width: 100%;">
                    <td><h4>Bundle Type</h4></td>
                    <td style="print-color-adjust:exact;"><h4>Totals</h4></td>
                </td>
        '''
        #Deposit summary query
        if DepositTotalIds: 
            sqlDepositTotals = q.QuerySql(SqlDepositTotals.format(searchsDate,searcheDate,DepositTotalIds))
            
            #loop through bundle summary
            TotalSource = 0.00
            
            for dt in sqlDepositTotals:
        
                #Column Totals
    
                TotalSource += dt.Total
                
                #Summary Details
                template += '''
                        <tr>
                            <td>{0}</td>
                            <td style="print-color-adjust:exact;">{1}</td>
                        </tr>
                    '''.format(dt.Source,checkNeg(dt.Total))
        
            template += '''
                <tr>
                    <td></td>
                    <td style="border-top: 1px solid black; print-color-adjust:exact;"><b>{0}</b></td>
                </tr></table></br>
            '''.format(checkNeg(TotalSource))
        else:
            template += '''<h3>Deposit Total Id(s) missing.  Open config to resolve</h3>'''
    
    if ShowFund == "yes":
    
        #Fund Summary    
        template +=  '''<h3>Funds</h3>
            <table id="trFundnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px;width: 100%;">
                <tr id="trFundnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px;width: 100%;">
                    <td><h4>Status</h4></td>
                    <td><h4>Fund Set</h4></td>
                    <td><h4>Fund</h4></td>
                    <td style="border-left: 1px dashed #dddddd;"><h4>General</h4></td>
                    <td><h4>Designated</h4></td>
                    <td class="subtotal-column" style="border-left: 1px dashed #dddddd;border-right: 1px dashed #dddddd;background-color: #f2f2f2;;print-color-adjust:exact;"><h4>Contribution<br />Total</h4></td>
                    <td><h4>NC</h4></td>
                    <td class="total-column" style="border-left: 1px solid #dddddd; border-right: 1px solid #dddddd; background-color: lightblue;print-color-adjust:exact;"><h4>Grand<br />Total</h4></td>
                </td>
        '''
        #fund summary query
        #print (SqlFunds.format(searchsDate,searcheDate,DateSearch,ExcludedFundIds,GeneralFundIds,NonContributionTypeIds))
        if FundSort == 'Fundset':
            FundSortBy = '''
                        CASE WHEN Fund = 'General Fund (1)' THEN 0 ELSE 1 END, 
                        FundSet,
                        Fund;
                        '''
        else:
            FundSortBy = '''
                        CASE WHEN Fund = 'General Fund (1)' THEN 0 ELSE 1 END, 
                        Fund;
                        '''
        
        sqlFunds = q.QuerySql(SqlFunds.format(searchsDate,searcheDate,DateSearch,ExcludedFundIds,GeneralFundIds,NonContributionTypeIds,FundSortBy))
        
        #loop through bundle summary
        TotalGeneralFund = 0.00
        TotalDesignatedFund = 0.00
        TotalNonContributionFund = 0.00
        TotalFund = 0.00
        GrandTotalFund = 0.00
        FundSet = ''
        FundSetName = ''
        FundSetStyle = ''
        
        for fs in sqlFunds:
            
                
            if FundSort == 'Fundset':
                if FundSet != fs.FundSet:
                    FundSetName = fs.FundSet
                    FundSet = fs.FundSet
                    FundSetStyle = ''
                else:
                    FundSetName = ''
                    FundSetStyle = '''style="position: relative; background: linear-gradient(to right, transparent 0%, lightblue 0%, lightblue 2px, transparent 2px); background-size: calc(100% - 20px) 2px; background-position: 20px center;"'''
            else: 
                FundSetName = fs.FundSet
                
            
            #Column Totals
            TotalGeneralFund += fs.General
            TotalDesignatedFund += fs.Designated
            TotalNonContributionFund += fs.NonContribution
            TotalFund += fs.General + fs.Designated
            GrandTotalFund += fs.General + fs.Designated + fs.NonContribution
            if fs.BundleStatus == 'Closed':
                BundleStatus = fs.BundleStatus
            else:
                BundleStatus = '<b>' + fs.BundleStatus + '</b>'
        

        
            #Summary Details
            template += '''
                    <tr>
                        <td>{0}</td>
                        <td {8}>{7}</td>
                        <td>{1}</td>
                        <td style="border-left: 1px dashed #dddddd;">{2}</td>
                        <td>{3}</td>
                        <td class="subtotal-column" style="border-left: 1px dashed #dddddd;border-right: 1px dashed #dddddd;background-color: #f2f2f2;;print-color-adjust:exact;">{5}</td>
                        <td>{4}</td>
                        <td class="total-column" style="border-left: 1px solid #dddddd;border-right: 1px solid #dddddd; background-color: lightblue;print-color-adjust:exact;">{6}</td>
                    </tr>
                '''.format(BundleStatus
                            ,fs.Fund
                            ,checkNeg(fs.General)
                            ,checkNeg(fs.Designated)
                            ,checkNeg(fs.NonContribution)
                            ,checkNeg(fs.General + fs.Designated)
                            ,checkNeg(fs.General + fs.Designated + fs.NonContribution)
                            ,FundSetName
                            ,FundSetStyle)
        
        template += '''
            <tr>
                <td></td>
                <td></td>
                <td></td>
                <td style="border-top: 1px solid #000000; border-left: 1px dashed #dddddd;"><b>{0}</b></td>
                <td style="border-top: 1px solid #000000;"><b>{1}</b></td>
                <td class="subtotal-column" style="border-left: 1px dotted #dddddd;border-right: 1px dotted #dddddd;border-top: 1px solid black;background-color: #f2f2f2;;print-color-adjust:exact;"><b>{3}</b></td>
                <td style="border-top: 1px solid #000000;"><b>{2}</b></td>
                <td class="total-column" style="border-top: 1px solid #000000; border-left: 1px solid #dddddd;border-right: 1px solid #dddddd; background-color: lightblue;print-color-adjust:exact;"><b>{4}</b></td>
            </tr></table></br>
        '''.format(checkNeg(TotalGeneralFund)
                ,checkNeg(TotalDesignatedFund)
                ,checkNeg(TotalNonContributionFund)
                ,checkNeg(TotalFund)
                ,checkNeg(GrandTotalFund))
        
        
    #bundle Details
    if ShowBundleDetails == "yes" and HideBundle != "yes":
        template += '''
            <hr><h2>Bundle Details</h2>
            '''
        
        #set Variables
        FundId = ''
        cType = ''
    
        sqlBundle = q.QuerySql(sqlBundle.format(searchsDate
                                                ,searcheDate
                                                ,DateSearch
                                                ,ExcludedFundIds
                                                ,GeneralFundIds
                                                ,NonContributionTypeIds))
        #loop through bundle type showing batch details for each
        for d in sqlBundle:
            if d.ContributionTypeId is None:
                ContributionTypeId = 'is Null'
            else:
                ContributionTypeId = '= ' + str(d.ContributionTypeId)
            
            #detail data pull
            HeaderTypeBundle = d.HeaderType.replace("'", "''") #d.HeaderType,
            sqlDetailsData = q.QuerySql(sqlDetails.format(searchsDate,
                                                            searcheDate,
                                                            HeaderTypeBundle, 
                                                            ExcludedFundIds,
                                                            GeneralFundIds,
                                                            DateSearch,
                                                            NonContributionTypeIds))
            
            #set table for each bundle
            template += '''
                <table>
                <tr>
                    <td><h3>{0} &nbsp</h3></td>
                </tr>'''.format(d.HeaderType,d.GeneralSum,d.DesignatedSum)
            template += '''</table>'''
            template += '''<table id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                <tr id="trnoborder" style="border: 1px solid #dddddd;text-align: left;padding: 10px;font-family: arial, sans-serif;border-collapse: collapse;max-width: 1000px; width: 100%;">
                    <td><h4>Status</h4></td>
                    <td><h4>Reference</h4></td>
                    <td><h4>Deposited</h4></td>
                    <td style="border-left: 1px dashed #dddddd;"><h4>General</h4></td>
                    <td><h4>Designated</h4></td>
                    <td class="subtotal-column" style="border-left: 1px dashed #dddddd; border-right: 1px dashed #dddddd; background-color: #f2f2f2;;print-color-adjust:exact;"><h4>Contribution<br>Total</h4></td>
                    <td><h4>NC</h4></td>
                    <td class="total-column" style="background-color: lightblue; border-left: 1px solid #dddddd; border-right: 1px solid #dddddd; print-color-adjust:exact;"><h4>Grand<br>Total</h4></td>
                </td>
            '''
            
            GrandTotalDetailGenDes = 0.00
            GrandTotalDetailGen = 0.00
            GrandTotalDetailDes = 0.00  
            GrandTotalDetailNon = 0.00
            GrandDetailTotal = 0.00
            
            #loop through each batch for the bundle
            for dd in sqlDetailsData:
        
                #Row Totals
                TotalDetailGenDes = dd.General + dd.Designated
                #Column Totals
                GrandTotalDetailGenDes += TotalDetailGenDes 
                GrandTotalDetailGen += dd.General
                GrandTotalDetailDes += dd.Designated
                GrandTotalDetailNon += dd.NonContribution
                GrandDetailTotal += dd.General + dd.Designated + dd.NonContribution
                
                FundDetails = ''
                if dd.SourceType is None:
                    cType = ''
                    cType = str(dd.ContributionType)
                else:
                    cType = str(dd.ContributionType) + ' (' + str(dd.SourceType) + ')'
        
                if ShowDesignatedDetails == 'yes':
                    #if dd.Designated != 0.00:
                    CombinedFunds = GeneralFundIds + ',' + ExcludedFundIds
                    sqlFundDetailsData = q.QuerySql(sqlFundDetails.format(dd.BundleHeaderId,CombinedFunds,DateSearch,GeneralFundIds,NonContributionTypeIds))
                    for fd in sqlFundDetailsData:
                        FundDetails += '''<tr>
                                            <td class="gray-text" colspan="3" style="color: gray; border-bottom: 1px dashed lightblue;">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp{0}</td>
                                            
                                            <td class="gray-text" style="color: gray; border-bottom: 1px dashed lightblue; border-left: 1px dashed #dddddd;">{1}</td>
                                            <td class="gray-text" style="color: gray; border-bottom: 1px dashed lightblue;">{2}</td>
                                            <td class="subtotal-column" style="border-left: 1px dashed #dddddd; border-right: 1px dashed #dddddd; border-bottom: 1px dashed lightblue; background-color: #f2f2f2;;print-color-adjust:exact;"></td>
                                            <td class="gray-text" style="color: gray; border-bottom: 1px dashed lightblue;">{3}</td>
                                            <td class="total-column" style="background-color: lightblue; border-left: 1px solid #dddddd; border-bottom: 1px dashed lightblue; border-right: 1px solid #dddddd; print-color-adjust:exact;"></td>
                                        </tr>
                                            '''.format(str(fd.Fund)
                                                      ,checkNeg(fd.General)
                                                      ,checkNeg(fd.Designated)
                                                      ,checkNeg(fd.NonContribution))

                if dd.BundleStatus != 'Closed':
                    BundleStatus = '<b>' + dd.BundleStatus + '</b>'
                else:
                    BundleStatus = dd.BundleStatus
                
                ReferenceId = '''<a href="/Batches/Detail/''' + str(dd.BundleHeaderId) + '''" target="_blank">''' + dd.ReferenceId + '''</a>&nbsp<a href="/PyScript/''' + BundleReportName + '''?p1=''' + str(dd.BundleHeaderId) + '''" target="_blank"><i class="fa fa-bar-chart" aria-hidden="true"></i></a>'''

                template += '''
                            <tr>
                                <td>{0}</td>
                                <td>{1}</td>
                                <td>{2}</td>
                                <td style="border-left: 1px dashed #dddddd;">{3}</td>
                                <td>{4}</td>
                                <td class="subtotal-column" style="border-left: 1px dashed #dddddd; border-right: 1px dashed #dddddd; background-color: #f2f2f2; print-color-adjust:exact;">{6}</td>
                                <td>{5}</td>
                                <td class="total-column" style="background-color: lightblue; border-right: 1px solid #dddddd; border-left: 1px solid #dddddd; border-bottom: 1px dashed lightblue; print-color-adjust:exact;">{7}</td>
                            </tr>
                            {8}    
                            '''.format(BundleStatus,ReferenceId,dd.DepositDate,checkNeg(dd.General),checkNeg(dd.Designated),checkNeg(dd.NonContribution),checkNeg(dd.General + dd.Designated),checkNeg(dd.General + dd.Designated + dd.NonContribution),FundDetails)

        
            template += '''
                <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td style="border-top: 1px solid black; border-left: 1px dashed #dddddd;"><b>{0}</b></td>
                    <td style="border-top: 1px solid black;"><b>{1}</b></td>
                    <td class="subtotal-column" style="border-left: 1px dashed #dddddd; border-right: 1px dashed #dddddd; border-top: 1px solid black;background-color: #f2f2f2;;print-color-adjust:exact;"><b>{3}</b></td>
                    <td style="border-top: 1px solid black;"><b>{2}</b></td>
                    <td class="total-column" style="border-top: 1px solid black; border-left: 1px solid #dddddd; border-right: 1px solid #dddddd; background-color: lightblue;print-color-adjust:exact;"><b>{4}</b></td>
                </tr>
            </table>
            '''.format(checkNeg(GrandTotalDetailGen)
                      ,checkNeg(GrandTotalDetailDes)
                      ,checkNeg(GrandTotalDetailNon)
                      ,checkNeg(GrandTotalDetailGenDes)
                      ,checkNeg(GrandDetailTotal))
        
            template += '''</table>'''
        
    NMReport = model.RenderTemplate(template) #render template and save to variable
    
    #send report to self if button is pressed
    if sendReport == 'y': 
        #Add Link Tracking
        NMReport += '{track}{tracklinks}<br />'
        
        #Set variables
        QueuedBy = model.UserPeopleId   # People ID of record the email should be queued by
        MailToQuery = model.UserPeopleId # '3134' 
    
        #Email
        model.Email(MailToQuery, QueuedBy, FromAddress, FromName, Subject, NMReport)
        
        #Notifiy User Report Sent
        print('<h3>Report Sent to Self</h3>') #Let people know report was sent
    else:
        #show button if sendReport <> y and print to screen the rendered template
        print '''<a href="?sendReport=y&sDate=''' + sDate + '''&eDate=''' + eDate + '''&ShowDesignatedDetails=''' + ShowDesignatedDetails + '''" target="_blank"><button type="button" class="no-print">Email to Self</button></a><hr>'''
        print(NMReport)
