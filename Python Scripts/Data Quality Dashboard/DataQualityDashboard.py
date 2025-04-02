#roles=admin

#--------------------------------------------------------------------
####REPORT INFORMATION####
#--------------------------------------------------------------------
#Tool to give an idea of quality of data
#
#Add all this code to a single Python Script
#1. Navigate to Admin ~ Advanced ~ Special Content ~ Python Scripts
#2. Select New Python Script File and Name the File
#3. Paste in all this code and Run
#Optional: Add to navigation menu

#--------------------------------------------------------------------
####USER CONFIG FIELDS
#--------------------------------------------------------------------
model.Header = "Data Quality Dashboard"

#--------------------------------------------------------------------
####START OF CODE.  No configuration should be needed beyond this point
#--------------------------------------------------------------------

import json

# Main query to calculate data quality metrics
sqlDataQuality = '''
WITH DataQualityMetrics AS (
    SELECT 
        CASE 
            WHEN p.IsDeceased = 1 AND p.ArchivedFlag = 1 THEN 'Deceased and Archived'
            WHEN p.IsDeceased = 1 THEN 'Deceased (Not Archived)'
            WHEN p.ArchivedFlag = 1 THEN 'Archived (Not Deceased)'
            ELSE 'Active'
        END AS RecordStatus,
        
        -- Filter to exclude archived records
        CASE WHEN p.ArchivedFlag = 0 THEN 1 ELSE 0 END AS IsActiveRecord,
        
        -- Core demographics
        CASE WHEN p.GenderId = 0 THEN 1 ELSE 0 END AS MissingGender,
        CASE WHEN p.BDate IS NULL AND (p.BirthMonth IS NULL OR p.BirthDay IS NULL OR p.BirthYear IS NULL) THEN 1 ELSE 0 END AS MissingBirthDate,
        CASE WHEN p.MaritalStatusId = 0 THEN 1 ELSE 0 END AS MissingMaritalStatus,
        CASE WHEN p.MemberStatusId IS NULL THEN 1 ELSE 0 END AS MissingMemberStatus,
        
        -- Church-specific data
        CASE WHEN p.BaptismStatusId = 0 THEN 1 ELSE 0 END AS MissingBaptismStatus,
        CASE WHEN p.CampusId IS NULL THEN 1 ELSE 0 END AS MissingCampus,
        
        -- Contact info 
        CASE WHEN (p.PrimaryAddress IS NULL OR p.PrimaryAddress = '') AND 
                  (p.AddressLineOne IS NULL OR p.AddressLineOne = '') THEN 1 ELSE 0 END AS MissingAddress,
        CASE WHEN (p.PrimaryCity IS NULL OR p.PrimaryCity = '') AND 
                  (p.CityName IS NULL OR p.CityName = '') THEN 1 ELSE 0 END AS MissingCity,
        CASE WHEN (p.PrimaryState IS NULL OR p.PrimaryState = '') AND 
                  (p.StateCode IS NULL OR p.StateCode = '') THEN 1 ELSE 0 END AS MissingState,
        CASE WHEN (p.PrimaryZip IS NULL OR p.PrimaryZip = '') AND 
                  (p.ZipCode IS NULL OR p.ZipCode = '') THEN 1 ELSE 0 END AS MissingZip,
        CASE WHEN p.BadAddressFlag = 1 OR p.PrimaryBadAddrFlag = 1 THEN 1 ELSE 0 END AS BadAddress,
        
        -- Phone info
        -- CASE WHEN (p.CellPhone IS NULL OR p.CellPhone = '') AND 
        --          (p.HomePhone IS NULL OR p.HomePhone = '') AND 
        --          (p.WorkPhone IS NULL OR p.WorkPhone = '') THEN 1 ELSE 0 END AS MissingPhone,
        CASE 
            WHEN p.Age < 13 THEN 
                CASE 
                    WHEN p.MemberStatusId = 0 THEN 1 
                    ELSE 0 
                END
            ELSE 
                CASE 
                    WHEN (p.CellPhone IS NULL OR p.CellPhone = '') AND 
                         (p.HomePhone IS NULL OR p.HomePhone = '') AND 
                         (p.WorkPhone IS NULL OR p.WorkPhone = '') THEN 1 
                    ELSE 0 
                END
        END AS MissingPhone,
        
        -- Email info
        --CASE WHEN (p.EmailAddress IS NULL OR p.EmailAddress = '') AND 
        --          (p.EmailAddress2 IS NULL OR p.EmailAddress2 = '') THEN 1 ELSE 0 END AS MissingEmail,
        CASE 
            WHEN p.Age < 13 THEN 
                CASE 
                    WHEN p.MemberStatusId = 0 THEN 1 
                    ELSE 0 
                END
            ELSE 
                CASE 
                    WHEN (p.EmailAddress IS NULL OR p.EmailAddress = '') AND 
                         (p.EmailAddress2 IS NULL OR p.EmailAddress2 = '') THEN 1 
                    ELSE 0 
                END
        END AS MissingEmail,
        
        
        -- Family info
        CASE WHEN p.FamilyId IS NULL THEN 1 ELSE 0 END AS MissingFamily,
        CASE WHEN p.PositionInFamilyId IS NULL THEN 1 ELSE 0 END AS MissingFamilyPosition,
        
        -- Photo
        CASE WHEN p.PictureId IS NULL THEN 1 ELSE 0 END AS MissingPhoto,
        
        -- Multiple of these flags ON indicates potential data quality issues
        (CASE WHEN p.DoNotMailFlag = 1 THEN 1 ELSE 0 END +
         CASE WHEN p.DoNotCallFlag = 1 THEN 1 ELSE 0 END +
         CASE WHEN p.DoNotVisitFlag = 1 THEN 1 ELSE 0 END +
         CASE WHEN p.DoNotPublishPhones = 1 THEN 1 ELSE 0 END) AS PrivacyRestrictions,
         
        -- For age demographics breakdown
        CASE 
            WHEN p.Age < 13 THEN 'Under 13'
            WHEN p.Age BETWEEN 13 AND 17 THEN '13-17'
            WHEN p.Age BETWEEN 18 AND 24 THEN '18-24'
            WHEN p.Age BETWEEN 25 AND 34 THEN '25-34'
            WHEN p.Age BETWEEN 35 AND 44 THEN '35-44'
            WHEN p.Age BETWEEN 45 AND 54 THEN '45-54'
            WHEN p.Age BETWEEN 55 AND 64 THEN '55-64'
            WHEN p.Age BETWEEN 65 AND 74 THEN '65-74'
            WHEN p.Age >= 75 THEN '75+'
            ELSE 'Unknown'
        END AS AgeRange,
        
        -- Organization involvement
        CASE WHEN p.BibleFellowshipClassId IS NULL THEN 1 ELSE 0 END AS MissingBibleClass,
        
        -- Data management status - useful to track when records might need review
        CASE WHEN p.ModifiedDate IS NULL THEN 1 ELSE 0 END AS NeverModified,
        CASE WHEN DATEDIFF(MONTH, p.ModifiedDate, GETDATE()) > 24 THEN 1 ELSE 0 END AS StaleRecord,
        
        1 AS PersonCount -- For calculating total records
        
    FROM People p
)

SELECT * FROM (
    -- Calculate aggregate statistics
    SELECT
        'OverallStats' AS MetricType,
        RecordStatus,
        COUNT(*) AS TotalRecords,
        SUM(MissingGender) AS MissingGender,
        SUM(MissingBirthDate) AS MissingBirthDate,
        SUM(MissingMaritalStatus) AS MissingMaritalStatus,
        SUM(MissingMemberStatus) AS MissingMemberStatus,
        SUM(MissingBaptismStatus) AS MissingBaptismStatus,
        SUM(MissingCampus) AS MissingCampus,
        SUM(MissingAddress) AS MissingAddress,
        SUM(MissingCity) AS MissingCity,
        SUM(MissingState) AS MissingState,
        SUM(MissingZip) AS MissingZip,
        SUM(BadAddress) AS BadAddress,
        SUM(MissingPhone) AS MissingPhone,
        SUM(MissingEmail) AS MissingEmail,
        SUM(MissingFamily) AS MissingFamily,
        SUM(MissingFamilyPosition) AS MissingFamilyPosition,
        SUM(MissingPhoto) AS MissingPhoto,
        SUM(MissingBibleClass) AS MissingBibleClass,
        SUM(NeverModified) AS NeverModified,
        SUM(StaleRecord) AS StaleRecord,
        
        -- Calculate percentages for key metrics
        CAST(SUM(MissingGender) * 100.0 / COUNT(*) AS DECIMAL(5,1)) AS PctMissingGender,
        CAST(SUM(MissingBirthDate) * 100.0 / COUNT(*) AS DECIMAL(5,1)) AS PctMissingBirthDate,
        CAST(SUM(MissingAddress) * 100.0 / COUNT(*) AS DECIMAL(5,1)) AS PctMissingAddress,
        CAST(SUM(MissingPhone) * 100.0 / COUNT(*) AS DECIMAL(5,1)) AS PctMissingPhone,
        CAST(SUM(MissingEmail) * 100.0 / COUNT(*) AS DECIMAL(5,1)) AS PctMissingEmail,
        CAST(SUM(MissingPhoto) * 100.0 / COUNT(*) AS DECIMAL(5,1)) AS PctMissingPhoto,
        
        -- Count records with privacy restrictions
        SUM(CASE WHEN PrivacyRestrictions > 0 THEN 1 ELSE 0 END) AS HasPrivacyRestrictions,
        
        -- Data completeness score (100% - average of missing critical data percentages)
        100 - (
            (CAST(SUM(MissingGender) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingBirthDate) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingAddress) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingPhone) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingEmail) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingMemberStatus) * 100.0 / COUNT(*) AS DECIMAL(5,1))) / 6
        ) AS DataCompletenessScore
        
    FROM DataQualityMetrics
    GROUP BY RecordStatus

    UNION ALL

    -- Age breakdown statistics
    SELECT
        'AgeBreakdown' AS MetricType,
        AgeRange AS RecordStatus,
        COUNT(*) AS TotalRecords,
        SUM(MissingGender) AS MissingGender,
        SUM(MissingBirthDate) AS MissingBirthDate,
        SUM(MissingMaritalStatus) AS MissingMaritalStatus,
        SUM(MissingMemberStatus) AS MissingMemberStatus,
        SUM(MissingBaptismStatus) AS MissingBaptismStatus,
        SUM(MissingCampus) AS MissingCampus,
        SUM(MissingAddress) AS MissingAddress,
        SUM(MissingCity) AS MissingCity,
        SUM(MissingState) AS MissingState,
        SUM(MissingZip) AS MissingZip,
        SUM(BadAddress) AS BadAddress,
        SUM(MissingPhone) AS MissingPhone,
        SUM(MissingEmail) AS MissingEmail,
        SUM(MissingFamily) AS MissingFamily,
        SUM(MissingFamilyPosition) AS MissingFamilyPosition,
        SUM(MissingPhoto) AS MissingPhoto,
        SUM(MissingBibleClass) AS MissingBibleClass,
        SUM(NeverModified) AS NeverModified,
        SUM(StaleRecord) AS StaleRecord,
        
        -- Calculate percentages for key metrics
        CAST(SUM(MissingGender) * 100.0 / NULLIF(COUNT(*), 0) AS DECIMAL(5,1)) AS PctMissingGender,
        CAST(SUM(MissingBirthDate) * 100.0 / NULLIF(COUNT(*), 0) AS DECIMAL(5,1)) AS PctMissingBirthDate,
        CAST(SUM(MissingAddress) * 100.0 / NULLIF(COUNT(*), 0) AS DECIMAL(5,1)) AS PctMissingAddress,
        CAST(SUM(MissingPhone) * 100.0 / NULLIF(COUNT(*), 0) AS DECIMAL(5,1)) AS PctMissingPhone,
        CAST(SUM(MissingEmail) * 100.0 / NULLIF(COUNT(*), 0) AS DECIMAL(5,1)) AS PctMissingEmail,
        CAST(SUM(MissingPhoto) * 100.0 / NULLIF(COUNT(*), 0) AS DECIMAL(5,1)) AS PctMissingPhoto,
        
        -- Count records with privacy restrictions
        SUM(CASE WHEN PrivacyRestrictions > 0 THEN 1 ELSE 0 END) AS HasPrivacyRestrictions,
        
        -- Data completeness score (100% - average of missing critical data percentages)
        100 - (
            (CAST(SUM(MissingGender) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingBirthDate) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingAddress) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingPhone) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingEmail) * 100.0 / COUNT(*) AS DECIMAL(5,1)) +
             CAST(SUM(MissingMemberStatus) * 100.0 / COUNT(*) AS DECIMAL(5,1))) / 6
        ) AS DataCompletenessScore
    FROM DataQualityMetrics
    --WHERE ArchivedFlag = 0 -- Only include non-archived records in age breakdown
    WHERE RecordStatus = 'Active'
    GROUP BY AgeRange
) AS combined_results
ORDER BY 
    MetricType,
    CASE 
        WHEN RecordStatus = 'Active' THEN 1
        WHEN RecordStatus = 'Dropped' THEN 2
        WHEN RecordStatus = 'Deceased' THEN 3
        WHEN RecordStatus = 'Archived' THEN 4
        WHEN RecordStatus = 'Unknown' THEN 5
        WHEN RecordStatus = 'Under 13' THEN 10
        WHEN RecordStatus = '13-17' THEN 11
        WHEN RecordStatus = '18-24' THEN 12
        WHEN RecordStatus = '25-34' THEN 13
        WHEN RecordStatus = '35-44' THEN 14
        WHEN RecordStatus = '45-54' THEN 15
        WHEN RecordStatus = '55-64' THEN 16
        WHEN RecordStatus = '65-74' THEN 17
        WHEN RecordStatus = '75+' THEN 18
        ELSE 19
    END
'''

# Query for specific people with missing data
sqlMissingDataPeople = '''
SELECT TOP 5000
    p.PeopleId,
    p.Name,
    p.Age,
    CASE 
        WHEN p.ArchivedFlag = 1 THEN 'Archived'
        WHEN p.IsDeceased = 1 THEN 'Deceased'
        ELSE 'Active'
    END AS RecordStatus,
    ms.Description AS MemberStatus,
    COALESCE(p.CampusId, 0) AS CampusId,
    
    -- Contact Info
    COALESCE(p.PrimaryAddress, p.AddressLineOne, '') AS Address,
    COALESCE(p.PrimaryCity, p.CityName, '') AS City,
    COALESCE(p.PrimaryState, p.StateCode, '') AS State,
    COALESCE(p.PrimaryZip, p.ZipCode, '') AS Zip,
    CASE WHEN p.Age >= 13 THEN COALESCE(p.CellPhone, '') END AS CellPhone,
    CASE WHEN p.Age >= 13 THEN COALESCE(p.HomePhone, '') END AS HomePhone,
    CASE WHEN p.Age >= 13 THEN COALESCE(p.EmailAddress, '') END AS Email,
    
    -- Missing Data Flags
    CASE WHEN p.GenderId = 0 THEN 1 ELSE 0 END AS MissingGender,
    CASE WHEN p.BDate IS NULL THEN 1 ELSE 0 END AS MissingBirthDate, --AND (p.BirthMonth IS NULL OR p.BirthDay IS NULL OR p.BirthYear IS NULL) 
    CASE WHEN p.MaritalStatusId = 0 THEN 1 ELSE 0 END AS MissingMaritalStatus,
    --CASE WHEN p.MemberStatusId IS NULL THEN 1 ELSE 0 END AS MissingMemberStatus,
    CASE WHEN p.BaptismStatusId = 0 THEN 1 ELSE 0 END AS MissingBaptismStatus,
    CASE WHEN p.CampusId IS NULL THEN 1 ELSE 0 END AS MissingCampus,
    CASE WHEN (p.PrimaryAddress IS NULL OR p.PrimaryAddress = '') AND 
              (p.AddressLineOne IS NULL OR p.AddressLineOne = '') THEN 1 ELSE 0 END AS MissingAddress,
    CASE WHEN p.Age >= 13 THEN
		CASE WHEN (p.CellPhone IS NULL OR p.CellPhone = '') AND 
              (p.HomePhone IS NULL OR p.HomePhone = '') AND 
              (p.WorkPhone IS NULL OR p.WorkPhone = '') THEN 1 ELSE 0 END 
			  END AS MissingPhone,
    CASE WHEN p.Age >= 13 THEN
		CASE WHEN (p.EmailAddress IS NULL OR p.EmailAddress = '') AND 
              (p.EmailAddress2 IS NULL OR p.EmailAddress2 = '') THEN 1 ELSE 0 END 
		END AS MissingEmail,
    CASE WHEN p.PictureId IS NULL THEN 1 ELSE 0 END AS MissingPhoto,
    
    -- Last modification
    p.ModifiedDate AS LastModified
    
FROM People p
LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId

WHERE 
    -- Only include records where some essential data is missing
    (
        p.GenderId IS NULL OR
        (p.BDate IS NULL AND (p.BirthMonth IS NULL OR p.BirthDay IS NULL OR p.BirthYear IS NULL)) OR
        p.MaritalStatusId IS NULL OR
        p.MemberStatusId IS NULL OR
        ((p.PrimaryAddress IS NULL OR p.PrimaryAddress = '') AND (p.AddressLineOne IS NULL OR p.AddressLineOne = '')) OR
        ((p.CellPhone IS NULL OR p.CellPhone = '') AND (p.HomePhone IS NULL OR p.HomePhone = '') AND (p.WorkPhone IS NULL OR p.WorkPhone = '')) OR
        ((p.EmailAddress IS NULL OR p.EmailAddress = '') AND (p.EmailAddress2 IS NULL OR p.EmailAddress2 = ''))
    )
    
    -- Focusing on active records by default
    AND (p.ArchivedFlag = 0 AND p.IsDeceased = 0)

ORDER BY p.ModifiedDate DESC
'''

# -------------------------------------------------------------------------
# Def Function
# -------------------------------------------------------------------------

# Near the beginning of the script, add a function to create SQL scripts
def create_sql_scripts():
    # Missing Email addresses
    email_sql = '''
    SELECT p.PeopleId, p.Name, p.Age
    FROM People p 
    WHERE (p.EmailAddress IS NULL OR p.EmailAddress = '') 
      AND (p.EmailAddress2 IS NULL OR p.EmailAddress2 = '')
      AND p.Age >= 13
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate DESC
    '''
    model.WriteContentSql("TPx_DQD-MissingEmailList", email_sql)
    
    # Missing Phone numbers
    phone_sql = '''
    SELECT p.PeopleId, p.Name, p.Age
    FROM People p 
    WHERE (p.CellPhone IS NULL OR p.CellPhone = '') 
      AND (p.HomePhone IS NULL OR p.HomePhone = '') 
      AND (p.WorkPhone IS NULL OR p.WorkPhone = '')
      AND p.Age >= 13
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate DESC
    '''
    model.WriteContentSql("TPx_DQD-MissingPhoneList", phone_sql)
    
    # Missing Addresses
    address_sql = '''
    SELECT p.PeopleId, p.Name, p.Age
    FROM People p 
    WHERE ((p.PrimaryAddress IS NULL OR p.PrimaryAddress = '') 
      AND (p.AddressLineOne IS NULL OR p.AddressLineOne = ''))
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate DESC
    '''
    model.WriteContentSql("TPx_DQD-MissingAddressList", address_sql)
    
    # Missing Gender
    gender_sql = '''
    SELECT p.PeopleId, p.Name, p.Age
    FROM People p 
    WHERE p.GenderId IS NULL OR p.GenderId = 0
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate DESC
    '''
    model.WriteContentSql("TPx_DQD-MissingGenderList", gender_sql)
    
    # Missing Birth Date
    birthdate_sql = '''
    SELECT p.PeopleId, p.Name, p.Age
    FROM People p 
    WHERE p.BDate IS NULL 
      AND (p.BirthMonth IS NULL OR p.BirthDay IS NULL OR p.BirthYear IS NULL)
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate DESC
    '''
    model.WriteContentSql("TPx_DQD-MissingBirthDateList", birthdate_sql)
    
    # Missing Photo
    photo_sql = '''
    SELECT p.PeopleId, p.Name, p.Age
    FROM People p 
    WHERE p.PictureId IS NULL
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate DESC
    '''
    model.WriteContentSql("TPx_DQD-MissingPhotoList", photo_sql)
    
    # Bad Addresses
    bad_address_sql = '''
    SELECT p.PeopleId, p.Name, p.Age, p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip
    FROM People p 
    WHERE (p.BadAddressFlag = 1 OR p.PrimaryBadAddrFlag = 1)
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate DESC
    '''
    model.WriteContentSql("TPx_DQD-BadAddressList", bad_address_sql)
    
    # Stale Records
    stale_records_sql = '''
    SELECT p.PeopleId, p.Name, p.Age, p.ModifiedDate
    FROM People p 
    WHERE DATEDIFF(MONTH, p.ModifiedDate, GETDATE()) > 24
      AND p.ArchivedFlag = 0 
      AND p.IsDeceased = 0
    ORDER BY p.ModifiedDate
    '''
    model.WriteContentSql("TPx_DQD-StaleRecordsList", stale_records_sql)

# Call this function early in your script
create_sql_scripts()

# -------------------------------------------------------------------------
# Dashboard HTML Template
# -------------------------------------------------------------------------

html_template = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Data Quality Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 100%; /* Changed from 1300px */
            margin: 0 auto;
            padding: 20px;
            overflow-x: hidden; /* Prevent horizontal scrolling */
            box-sizing: border-box; /* Add this to include padding in width calculation */
        }
        
        h1, h2, h3, h4 {
            color: #2c3e50;
        }

        /* Main dashboard title */
        h1 {
            font-size: 2.2rem;
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 20px;
        }

        /* Section headings should be bigger too */
        h2 {
            font-size: 1.8rem !important;
            font-weight: 600 !important;
            color: #2c3e50 !important;
            margin-top: 30px !important;
            margin-bottom: 25px !important;
            padding-bottom: 15px !important;
            border-bottom: 2px solid #eee !important;
        }

        /* Secondary headings and card titles */
        h3, .metric-label {
            font-size: 1.1rem;
            font-weight: 600;
            color: #34495e;
            margin-top: 0;
            margin-bottom: 10px;
        }

        .container {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            margin-bottom: 30px;
        }

        .container-fluid {
            padding-left: 0;
            padding-right: 0;
            overflow-x: hidden;
        }
        
        .metric-card {
            background-color: #fff;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            padding: 20px;
            flex: 1;
            min-width: 200px;
            text-align: center;
        }

        /* Main metric card titles */
        .metric-card .metric-label {
            font-size: 1.3rem !important; /* Increase font size significantly */
            font-weight: 600 !important;  /* Make font bolder */
            color: #2c3e50 !important;    /* Darker color for better contrast */
            margin-bottom: 12px !important; /* Add more space below */
            text-align: center !important; /* Center align for better appearance */
            line-height: 1.4 !important;  /* Better line height */
        }
        
        /* The metric value itself */
        .metric-value {
            font-size: 2.8rem !important; /* Make values bigger */
            font-weight: 700 !important;  /* Make them bolder */
            color: #2c3e50 !important;    /* Consistent color */
            margin: 15px 0 !important;    /* Good spacing */
            line-height: 1.2 !important;  /* Better line height */
        }
        
        .metric-label {
            font-size: 1rem;
            color: #34495e;
            font-weight: 600;
            margin-bottom: 5px;
        }
        
        /* Fixed height chart containers */
        .chart-container {
            width: 100%;
            height: 400px; /* Fixed height */
            margin-bottom: 30px;
            background-color: #fff;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            padding: 20px;
            box-sizing: border-box;
            overflow: hidden; /* Prevent content from overflowing */
        }
        
        .chart-row {
            display: flex;
            flex-wrap: wrap; /* Allow wrapping on smaller screens */
            gap: 20px;
            margin-bottom: 30px;
        }
        
        /* Fix for the chart containers */
        .chart-container, .chart-row, .chart-card {
            max-width: 100%;
            box-sizing: border-box;
        }        

        .chart-card {
            background-color: #fff;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            padding: 20px;
            flex: 1;
            min-width: 300px;
            height: 350px; /* Fixed height */
            position: relative;
            overflow: hidden; /* Hide overflowing content */
            padding-bottom: 5px;
            margin-bottom: 5px;
        }
        
        /* Add important to chart titles as well */
        .chart-card h3 {
            font-size: 1.35rem !important;
            font-weight: 600 !important;
            color: #2c3e50 !important;
            margin-top: 0 !important;
            margin-bottom: 15px !important;
            text-align: center !important;
        }
        
        /* Make charts respect container height */
        .chart-card canvas {
            max-height: 280px !important; /* Allow room for title */
            width: 100% !important;
        }
        
        .table-container {
            position: relative;
            max-height: 1000px;
            overflow-y: auto;
            overflow-x: auto;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            padding: 0; /* Remove padding */
            margin-bottom: 30px;
        }

        /* Table section titles */
        .table-container h3 {
            font-size: 1.5rem !important;
            margin-bottom: 20px !important;
            font-weight: 600 !important;
        }
        
        table {
            width: 100%;
            max-width: 100%;
            border-collapse: collapse;
            table-layout: fixed; /* Fixed layout for better performance */
        }

        /* Make table headers more prominent */
        table th {
            font-size: 1.1rem !important;
            font-weight: 600 !important;
            color: #34495e !important;
            padding: 15px !important;
        }
        
        th, td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #e0e0e0;
            white-space: nowrap; /* Prevent wrapping */
        }
        
        th {
            background-color: #f5f7fa;
            font-weight: 600;
            position: sticky;
            top: 0; /* Sticky headers */
            z-index: 10;
        }
        
        tr:hover {
            background-color: #f8f9fa;
        }
        
        /* Make the score rating more prominent */
        .score-label {
            font-size: 1.6rem !important;
            margin-top: 15px !important;
            font-weight: 600 !important;
        }
        
        .score-container {
            position: relative;
            width: 200px;
            height: 200px;
            margin: 0 auto;
        }
        
        .badge {
            display: inline-block;
            padding: 3px 7px;
            border-radius: 50px;
            font-size: 12px;
            font-weight: 600;
            margin-right: 5px;
            margin-bottom: 3px;
        }
        
        .badge-warning {
            background-color: #ff9800;
            color: white;
        }
        
        .badge-danger {
            background-color: #f44336;
            color: white;
        }
        
        .badge-info {
            background-color: #2196f3;
            color: white;
        }
        
        .badge-success {
            background-color: #4caf50;
            color: white;
        }
        
        /* Tabs */
        .tabs {
            display: flex;
            margin-bottom: 20px;
            flex-wrap: wrap; /* Allow wrapping on mobile */
        }
        
        .tab {
            padding: 10px 20px;
            cursor: pointer;
            border: 1px solid #ddd;
            border-bottom: none;
            border-radius: 5px 5px 0 0;
            background-color: #f5f7fa;
            white-space: nowrap; /* Prevent tab text wrapping */
        }
        
        .tab.active {
            background-color: #fff;
            border-bottom: 2px solid #3498db;
            font-weight: bold;
        }
        
        .tab-content {
            display: none;
        }
        
        .tab-content.active {
            display: block;
        }
        
        /* Fix for data quality statistics table */
        #status-table th,
        #age-table th {
            font-size: 1rem;
            padding: 15px 10px;
        }
        
        .sort-btn {
            border: none;
            background: none;
            padding: 0;
            margin-left: 3px;
            cursor: pointer;
            font-size: 0.8rem;
            color: #7f8c8d;
        }
        
        /* Filter controls for problem records */
        .filter-controls {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            margin-bottom: 20px;
            align-items: center;
        }
        
        .filter-controls select {
            padding: 8px;
            border-radius: 4px;
            border: 1px solid #ddd;
        }
        
        /* The help text below */
        .help-text {
            font-size: 1rem !important;   /* Larger help text */
            color: #34495e !important;    /* Darker for readability */
            line-height: 1.4 !important;  /* Better line height */
        }
        
        /* Filter labels */
        .filter-controls label {
            font-size: 1rem;
            font-weight: 600;
        }
        
        /* Loading indicator */
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }
        
        .spinner {
            border: 4px solid rgba(0, 0, 0, 0.1);
            width: 36px;
            height: 36px;
            border-radius: 50%;
            border-left-color: #3498db;
            animation: spin 1s linear infinite;
            margin: 0 auto;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        /* Pagination controls */
        .pagination {
            display: flex;
            justify-content: center;
            margin-top: 20px;
            gap: 5px;
            flex-wrap: wrap;
        }
        
        .pagination button {
            padding: 8px 15px;
            background-color: #f5f7fa;
            border: 1px solid #ddd;
            cursor: pointer;
            border-radius: 4px;
        }
        
        .pagination button.active {
            background-color: #3498db;
            color: white;
            border-color: #3498db;
        }
        
        /* Mobile optimizations */
        @media (max-width: 768px) {
            .chart-row {
                flex-direction: column;
            }
            
            .chart-card {
                min-width: 100%;
            }
            
            .metric-card {
                min-width: 100%;
            }
            
            .table-container {
                padding: 10px;
            }
            
            th, td {
                padding: 8px 10px;
            }
        }

        #problem-records-table {
            border-collapse: separate;
            border-spacing: 0;
            width: 100%;
        }

        /* Properly sticky headers */
        #problem-records-table thead {
            position: sticky;
            top: 0;
            z-index: 1000; /* Very high z-index */
        }
        
        #problem-records-table th {
            background-color: #f5f7fa;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 1000; /* Very high z-index */
            border-bottom: 2px solid #3498db;
            padding: 12px 15px;
            box-shadow: 0 2px 2px -1px rgba(0, 0, 0, 0.1);
        }
        
        /* Add a wrapper inside the table container for padding */
        .table-inner-container {
            padding: 20px;
            padding-top: 0; /* No top padding */
        }
        
        
        /* Fix for column headers */
        #problem-records-table th:nth-child(3) {
            white-space: nowrap;
            min-width: 140px;
        }
        
        /* Make Missing Data column wider and allow wrapping */
        #problem-records-table td:nth-child(3) {
            white-space: normal; /* Allow wrapping */
            min-width: 160px;
            max-width: 220px;
        }
        
        /* Ensure Last Modified column has enough space */
        #problem-records-table th:nth-child(4),
        #problem-records-table td:nth-child(4) {
            white-space: nowrap;
            min-width: 150px;
            width: 150px;
        }
        
        /* Badges should display properly */
        #problem-records-table td:nth-child(3) {
            display: flex;
            flex-wrap: wrap;
            gap: 4px;
            padding-top: 8px;
            padding-bottom: 8px;
        }
        
        /* Individual badges */
        .badge {
            display: inline-block;
            margin-right: 3px;
            margin-bottom: 3px;
        }

        /* Pop-up */
        .info-icon {
            cursor: pointer;
            color: #3498db;
            margin-left: 5px;
            font-size: 1rem;
        }
        
        .tooltip-popup {
            position: fixed;
            z-index: 1000;
            background-color: white;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 15px;
            max-width: 400px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            display: none;
        }
    </style>
</head>
<body>
    <h1>
      Data Quality Dashboard
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
        <!-- Text portion - TP -->
        <text x="100" y="120" font-family="Arial, sans-serif" font-weight="bold" font-size="60" fill="#333333">TP</text>
        
        <!-- Circular element -->
        <g transform="translate(190, 107)">
          <!-- Outer circle -->
          <circle cx="0" cy="0" r="13.5" fill="#0099FF"/>
          
          <!-- White middle circle -->
          <circle cx="0" cy="0" r="10.5" fill="white"/>
          
          <!-- Inner circle -->
          <circle cx="0" cy="0" r="7.5" fill="#0099FF"/>
          
          <!-- X crossing through the circles -->
          <path d="M-9 -9 L9 9 M-9 9 L9 -9" stroke="white" stroke-width="1.8" stroke-linecap="round"/>
        </g>
        
        <!-- Single "i" letter to the right -->
        <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">i</text>
      </svg>
    </h1>

    
    <div class="tabs">
        <div class="tab active" data-tab="overview">Overview</div>
        <div class="tab" data-tab="records">Problem Records</div>
        <div class="tab" data-tab="actions">Recommended Actions</div>
    </div>

    <!-- Overview Tab -->
    <div id="overview" class="tab-content active">
        <h2>Data Health Overview</h2>
        
        <!-- Top level metrics -->
        <div class="container">
            <div class="metric-card">
                <div class="metric-label">Total Active Records</div>
                <div class="metric-value" id="total-records">{{active_count}}</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Data Completeness Score 
                    <span class="info-icon" onclick="showDataCompletenessTooltip(event)">ℹ️</span>
                </div>
                <div class="metric-value" id="completeness-score">{{data_completeness_score}}%</div>
                <div id="score-rating" class="score-label">
                    {{score_rating}}
                </div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Missing Critical Data</div>
                <div class="metric-value" id="missing-critical">{{missing_critical_count}}</div>
                <div class="help-text">Records missing address, phone, or email</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Stale Records</div>
                <div class="metric-value" id="stale-records">{{stale_records}}</div>
                <div class="help-text">Not updated in over 24 months</div>
            </div>
        </div>

        <!-- Charts for data completeness -->
        <div class="chart-row">
            <div class="chart-card">
                <h3>Records by Status</h3>
                <canvas id="status-chart"></canvas>
            </div>
            
            <div class="chart-card">
                <h3>Missing Data by Category</h3>
                <canvas id="missing-data-chart"></canvas>
            </div>
        </div>
        
        <div class="chart-row">
            <div class="chart-card">
                <h3>Data by Age Group</h3>
                <canvas id="age-data-chart"></canvas>
            </div>
            
            <div class="chart-card">
                <h3>Contact Information</h3>
                <canvas id="contact-chart"></canvas>
            </div>
        </div>
        
        <!-- Data quality by status table -->
        <div class="table-container">
            <h3>Data Quality Statistics by Record Status</h3>
            <table id="status-table">
                <thead>
                    <tr>
                        <th>Status</th>
                        <th>Count</th>
                        <th>Missing Gender <button class="sort-btn" onclick="sortTable('status-table', 2)">↕</button></th>
                        <th>Missing Birth Date <button class="sort-btn" onclick="sortTable('status-table', 3)">↕</button></th>
                        <th>Missing Address <button class="sort-btn" onclick="sortTable('status-table', 4)">↕</button></th>
                        <th>Missing Phone <span class="info-icon" onclick="showPhoneEmailTooltip(event, 'phone')">ℹ️</span> <button class="sort-btn" onclick="sortTable('status-table', 5)">↕</button></th>
                        <th>Missing Email <span class="info-icon" onclick="showPhoneEmailTooltip(event, 'email')">ℹ️</span> <button class="sort-btn" onclick="sortTable('status-table', 6)">↕</button></th>
                        <th>Data Score <button class="sort-btn" onclick="sortTable('status-table', 7, true)">↕</button></th>
                    </tr>
                </thead>
                <tbody>
                    {{status_table_rows}}
                </tbody>
            </table>
        </div>
        
        <!-- Data quality by age group table - with scrolling -->
        <div class="table-container">
            <h3>Data Quality Statistics by Age Group (Excludes Archive)</h3>
            <table id="age-table">
                <thead>
                    <tr>
                        <th>Age Group</th>
                        <th>Count</th>
                        <th>Missing Gender <button class="sort-btn" onclick="sortTable('age-table', 2)">↕</button></th>
                        <th>Missing Address <button class="sort-btn" onclick="sortTable('age-table', 4)">↕</button></th>
                        <th>Missing Phone <span class="info-icon" onclick="showPhoneEmailTooltip(event, 'phone')">ℹ️</span> <button class="sort-btn" onclick="sortTable('age-table', 5)">↕</button></th>
                        <th>Missing Email <span class="info-icon" onclick="showPhoneEmailTooltip(event, 'email')">ℹ️</span> <button class="sort-btn" onclick="sortTable('age-table', 6)">↕</button></th>
                        <th>Data Score <button class="sort-btn" onclick="sortTable('age-table', 7, true)">↕</button></th>
                    </tr>
                </thead>
                <tbody>
                    {{age_table_rows}}
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- Problem Records Tab -->
    <div id="records" class="tab-content">
        <h2>Problem Records</h2>
        <p>This table shows records with missing critical data that may need attention. Click on a person ID to view their record.</p>
        
        <!-- Filter controls -->
        <div class="filter-controls">
            <div>
                <label for="issue-filter">Filter by issue:</label>
                <select id="issue-filter" onchange="filterProblemRecords()">
                    <option value="all">All Issues</option>
                    <option value="missing-email">Missing Email (13+)</option>
                    <option value="missing-phone">Missing Phone (13+)</option>
                    <option value="missing-address">Missing Address</option>
                    <option value="missing-gender">Missing Gender</option>
                    <option value="missing-birthdate">Missing Birth Date</option>
                </select>
            </div>
            
            <div>
                <label for="records-per-page">Records per page:</label>
                <select id="records-per-page" onchange="changeRecordsPerPage()">
                    <option value="25">25</option>
                    <option value="50">50</option>
                    <option value="100">100</option>
                </select>
            </div>
            
            <div id="records-count"></div>
        </div>
        
        <!-- Loading indicator -->
        <div id="problem-records-loading" class="loading">
            <div class="spinner"></div>
            <p>Loading records...</p>
        </div>
        
        <div class="table-container">
            <table id="problem-records-table">
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>Age</th>
                        <th>Missing Data</th>
                        <th>Last Modified</th>
                    </tr>
                </thead>
                <tbody>
                    {{problem_records_rows}}
                </tbody>
            </table>
            
            <div class="pagination" id="records-pagination"></div>
        </div>
    </div>
    
    <!-- Recommended Actions Tab -->
    <div id="actions" class="tab-content">
        <h2>Recommended Data Quality Actions</h2>
        
        <div class="container">
            <div class="metric-card">
                <div class="metric-label">Overall Data Health</div>
                <div class="metric-value">{{health_rating}}</div>
                <div class="help-text">Based on completeness score</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Critical Action Items</div>
                <div class="metric-value">{{critical_actions}}</div>
                <div class="help-text">High-priority data fixes</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Cleanup Opportunities</div>
                <div class="metric-value">{{cleanup_opportunities}}</div>
                <div class="help-text">Potential data improvements</div>
            </div>
        </div>
        
        <div class="table-container">
            <h3>Recommended Action Items</h3>
            <table>
                <thead>
                    <tr>
                        <th>Priority</th>
                        <th>Task</th>
                        <th>Impact</th>
                        <th>Records Affected</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody>
                    {{action_items}}
                </tbody>
            </table>
        </div>
        
        <div class="table-container">
            <h3>Data Collection Opportunities</h3>
            <p>Consider implementing these strategies to improve your data collection process:</p>
            <ul>
                <li><strong>Update member information forms</strong> - Ensure all critical fields are required during data collection</li>
                <li><strong>Annual data verification</strong> - Ask members to verify their information yearly</li>
                <li><strong>Mobile app updates</strong> - Allow members to update their own information easily</li>
                <li><strong>Photo collection events</strong> - Schedule specific times for taking member photos</li>
                <li><strong>Email/SMS verification</strong> - Implement automated verification of contact details</li>
                <li><strong>Data quality reports</strong> - Run this dashboard monthly to track improvements</li>
            </ul>
        </div>
    </div>
</body>
</html>
    
    <script>
    // Dashboard initialization and global variables
    let problemRecordsLoaded = false;
    let currentPage = 1;
    let recordsPerPage = 25;
    let problemRecordsData = [];
    let filteredRecords = [];
    
    // Defer chart initialization to improve initial page load
    let statusChart, missingDataChart, ageDataChart, contactChart;
    let chartsInitialized = false;
    
    // Tab functionality with lazy loading
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => {
            // Remove active class from all tabs and content
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            // Add active class to clicked tab and corresponding content
            tab.classList.add('active');
            document.getElementById(tab.dataset.tab).classList.add('active');
            
            // Load data for the tab if needed
            if (tab.dataset.tab === 'overview' && !chartsInitialized) {
                initializeCharts();
            } else if (tab.dataset.tab === 'records' && !problemRecordsLoaded) {
                loadProblemRecords();
            }
        });
    });
    
    // Initialize charts only when needed
    function initializeCharts() {
        if (chartsInitialized) return;
        
        console.log("Initializing charts...");
        
        // Register Chart.js plugins
        Chart.register(ChartDataLabels);
        
        // Set chart defaults to improve performance
        Chart.defaults.font.size = 12;
        Chart.defaults.animation.duration = 500;
        Chart.defaults.plugins.tooltip.enabled = true;
        
        // Optimized status chart - focuses on key statuses, consolidates small ones
        initStatusChart();
        
        // Optimized missing data chart - only shows critical fields
        initMissingDataChart();
        
        // Age group chart with simplified visualization
        initAgeDataChart();
        
        // Contact information chart
        initContactChart();
        
        chartsInitialized = true;
        console.log("Charts initialized");
    }
    
    function initStatusChart() {
        const statusCtx = document.getElementById('status-chart').getContext('2d');
        const statusLabels = {{status_labels}};
        const statusData = {{status_data}};
        
        // Consolidate small categories (less than 3% of total)
        const total = statusData.reduce((a, b) => a + b, 0);
        const threshold = total * 0.03;
        
        const consolidatedLabels = [];
        const consolidatedData = [];
        let otherTotal = 0;
        
        statusData.forEach((value, index) => {
            if (value >= threshold) {
                consolidatedLabels.push(statusLabels[index]);
                consolidatedData.push(value);
            } else {
                otherTotal += value;
            }
        });
        
        if (otherTotal > 0) {
            consolidatedLabels.push('Other');
            consolidatedData.push(otherTotal);
        }
        
        statusChart = new Chart(statusCtx, {
            type: 'pie',
            data: {
                labels: consolidatedLabels,
                datasets: [{
                    data: consolidatedData,
                    backgroundColor: [
                        '#4caf50',
                        '#ff9800',
                        '#9e9e9e',
                        '#e0e0e0',
                        '#607d8b'
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        color: '#fff',
                        font: { weight: 'bold' },
                        formatter: (value, ctx) => {
                            const label = ctx.chart.data.labels[ctx.dataIndex];
                            const percentage = Math.round((value / total) * 100);
                            return `${percentage}%`;
                        }
                    },
                    legend: {
                        position: 'right',
                        labels: { boxWidth: 15 }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const percentage = Math.round((context.raw / total) * 100);
                                return `${context.label}: ${context.raw.toLocaleString()} (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });
    }

    const commonChartOptions = {
        responsive: true,
        maintainAspectRatio: false,  // This is crucial!
        plugins: {
            legend: {
                position: 'top',
                labels: {
                    boxWidth: 12,
                    font: {
                        size: 11
                    }
                }
            }
        }
    };
    
    function initMissingDataChart() {
        try {
            const missingDataCtx = document.getElementById('missing-data-chart').getContext('2d');
            
            // Parse the JSON string back into an array
            const missingDataPercentages = {{missing_data_percentages}};
            
            // Validate data
            if (!Array.isArray(missingDataPercentages) || missingDataPercentages.length === 0) {
                console.warn("No missing data percentages available");
                return;
            }
            
            // Create labels with additional info
            const missingDataLabels = [
                'Gender', 
                'Birth Date', 
                'Address', 
                'Phone', 
                'Email', 
                'Photo'
            ];
            
            missingDataChart = new Chart(missingDataCtx, {
                type: 'bar',
                data: {
                    labels: missingDataLabels,
                    datasets: [{
                        label: 'Missing Data Percentage',
                        data: missingDataPercentages,
                        backgroundColor: [
                            'rgba(255, 99, 132, 0.6)',   // Gender
                            'rgba(54, 162, 235, 0.6)',   // Birth Date
                            'rgba(255, 206, 86, 0.6)',   // Address
                            'rgba(75, 192, 192, 0.6)',   // Phone
                            'rgba(153, 102, 255, 0.6)',  // Email
                            'rgba(255, 159, 64, 0.6)'    // Photo
                        ],
                        borderColor: [
                            'rgba(255, 99, 132, 1)',
                            'rgba(54, 162, 235, 1)',
                            'rgba(255, 206, 86, 1)',
                            'rgba(75, 192, 192, 1)',
                            'rgba(153, 102, 255, 1)',
                            'rgba(255, 159, 64, 1)'
                        ],
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: 'Percentage Missing (%)'
                            }
                        }
                    },
                    plugins: {
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    return context.parsed.y + '% Missing';
                                }
                            }
                        },
                        datalabels: {
                            display: true,
                            color: 'black',
                            font: {
                                weight: 'bold'
                            },
                            formatter: function(value) {
                                return value + '%';
                            }
                        }
                    }
                },
                plugins: [ChartDataLabels]
            });
    
            // Add event listener for the "Data Completeness Score" info icon
            const dataCompletenessIcon = document.querySelector('.metric-label .info-icon');
            if (dataCompletenessIcon) {
                dataCompletenessIcon.addEventListener('click', function(event) {
                    showDataCompletenessTooltip(event);
                });
            }
    
            // Add event listeners for the "Phone/Email" info icons in the chart
            const phoneEmailIcons = document.querySelectorAll('#missing-data-chart + .info-icon');
            phoneEmailIcons.forEach(icon => {
                icon.addEventListener('click', function(event) {
                    const type = this.getAttribute('data-type');
                    showPhoneEmailTooltip(event, type);
                });
            });
    
        } catch (error) {
            console.error("Error initializing missing data chart:", error);
        }
    }

    
    function initAgeDataChart() {
        const ageDataCtx = document.getElementById('age-data-chart').getContext('2d');
        
        // Parse age group data from the template
        const ageLabels = {{age_group_labels}};
        const ageScores = {{age_group_scores}};
        const ageCounts = {{age_group_counts}}; // You'll need to add this in your Python code
        
        ageDataChart = new Chart(ageDataCtx, {
            type: 'bar',
            data: {
                labels: ageLabels,
                datasets: [{
                    type: 'line',
                    label: 'Data Completeness Score',
                    data: ageScores,
                    borderColor: '#e74c3c',
                    backgroundColor: 'rgba(231, 76, 60, 0.2)',
                    borderWidth: 2,
                    tension: 0.1,
                    fill: false,
                    yAxisID: 'y1'
                }, {
                    type: 'bar',
                    label: 'Number of Records',
                    data: ageCounts,
                    backgroundColor: '#3498db',
                    borderWidth: 1,
                    yAxisID: 'y'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        display: false
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                if (context.dataset.type === 'line') {
                                    return `Score: ${context.raw}%`;
                                } else {
                                    return `Records: ${context.raw.toLocaleString()}`;
                                }
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Number of Records'
                        },
                        type: 'linear',
                        position: 'left'
                    },
                    y1: {
                        beginAtZero: true,
                        max: 100,
                        title: {
                            display: true,
                            text: 'Data Completeness Score (%)'
                        },
                        type: 'linear',
                        position: 'right',
                        grid: {
                            drawOnChartArea: false
                        }
                    }
                }
            }
        });
    }
    
    function initContactChart() {
        const contactCtx = document.getElementById('contact-chart').getContext('2d');
        
        // Parse contact data from the template
        const contactData = {{contact_data}};
        const contactLabels = [
            'Has All Contact Info', 
            'Missing Email Only', 
            'Missing Phone Only', 
            'Missing Address Only', 
            'Multiple Missing'
        ];
        
        contactChart = new Chart(contactCtx, {
            type: 'pie',
            data: {
                labels: contactLabels,
                datasets: [{
                    data: contactData,
                    backgroundColor: [
                        '#2ecc71',   // All contact info
                        '#f39c12',   // Missing email
                        '#3498db',   // Missing phone
                        '#e74c3c',   // Missing address
                        '#9b59b6'    // Multiple missing
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        color: '#fff',
                        font: { weight: 'bold' },
                        formatter: (value, ctx) => {
                            const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = Math.round((value / total) * 100);
                            if (percentage < 5) return '';
                            return `${percentage}%`;
                        }
                    },
                    legend: {
                        position: 'right',
                        labels: { 
                            boxWidth: 15,
                            generateLabels: function(chart) {
                                const data = chart.data;
                                const total = data.datasets[0].data.reduce((a, b) => a + b, 0);
                                return data.labels.map((label, index) => {
                                    const value = data.datasets[0].data[index];
                                    const percentage = Math.round((value / total) * 100);
                                    return {
                                        text: `${label} (${percentage}%)`,
                                        fillStyle: data.datasets[0].backgroundColor[index]
                                    };
                                });
                            }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const value = context.raw;
                                const percentage = Math.round((value / total) * 100);
                                return `${context.label}: ${value} (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });
    }
    
    // Problem records pagination and filtering
    function loadProblemRecords() {
        document.getElementById('problem-records-loading').style.display = 'block';
        
        if (problemRecordsLoaded) {
            document.getElementById('problem-records-loading').style.display = 'none';
            return;
        }
        
        // When using pre-loaded rows, immediately process the data
        const preLoadedRecords = document.querySelectorAll('#problem-records-table tbody tr');
        if (preLoadedRecords.length > 0) {
            console.log("Found " + preLoadedRecords.length + " pre-loaded problem records");
            
            // Process the pre-loaded records
            problemRecordsData = [];
            
            preLoadedRecords.forEach(record => {
                const cells = record.querySelectorAll('td');
                if (cells.length >= 4) {
                    // Extract the person ID from the link in the Name column
                    const nameCell = cells[0];
                    const nameLink = nameCell.querySelector('a');
                    const personId = nameLink ? nameLink.getAttribute('href').split('/').pop() : '';
                    const name = nameLink ? nameLink.textContent.trim() : nameCell.textContent.trim();
                    
                    problemRecordsData.push({
                        PersonId: personId,
                        Name: name,
                        Age: cells[1].textContent.trim(),
                        MissingDataTags: cells[2].innerHTML,
                        LastModified: cells[3].textContent.trim(),
                        MissingGender: cells[2].innerHTML.includes('Gender') ? 1 : 0,
                        MissingBirthDate: cells[2].innerHTML.includes('Birth Date') ? 1 : 0,
                        MissingAddress: cells[2].innerHTML.includes('Address') ? 1 : 0,
                        MissingPhone: cells[2].innerHTML.includes('Phone') ? 1 : 0,
                        MissingEmail: cells[2].innerHTML.includes('Email') ? 1 : 0,
                        MissingPhoto: cells[2].innerHTML.includes('Photo') ? 1 : 0
                    });
                }
            });
            
            filteredRecords = [...problemRecordsData];
            
            // Clear the original table data to replace with paginated data
            const tableBody = document.querySelector('#problem-records-table tbody');
            tableBody.innerHTML = '';
            
            generatePagination();
            displayRecordsPage(1);
            
            document.getElementById('problem-records-loading').style.display = 'none';
            problemRecordsLoaded = true;
            
            // Add filtering functionality
            document.getElementById('issue-filter').addEventListener('change', filterProblemRecords);
            document.getElementById('records-per-page').addEventListener('change', changeRecordsPerPage);
            
            // Display record count
            const recordCountElement = document.getElementById('records-count');
            if (recordCountElement) {
                recordCountElement.textContent = `Showing 1-${Math.min(recordsPerPage, problemRecordsData.length)} of ${problemRecordsData.length} records`;
            }
            
            return; // Exit early, no need for the timeout
        }
        
        // If no pre-loaded records, display a message
        document.getElementById('problem-records-loading').style.display = 'none';
        
        const tableBody = document.querySelector('#problem-records-table tbody');
        tableBody.innerHTML = '<tr><td colspan="4" style="text-align: center;">No problem records found</td></tr>';
        
        const recordCountElement = document.getElementById('records-count');
        if (recordCountElement) {
            recordCountElement.textContent = "0 records found";
        }
        
        problemRecordsLoaded = true;
    }
    
    function generatePagination() {
        const paginationElement = document.getElementById('records-pagination');
        paginationElement.innerHTML = '';
        
        const totalPages = Math.ceil(filteredRecords.length / recordsPerPage);
        
        // Previous button
        if (totalPages > 1) {
            const prevButton = document.createElement('button');
            prevButton.textContent = '←';
            prevButton.addEventListener('click', () => {
                if (currentPage > 1) {
                    displayRecordsPage(currentPage - 1);
                }
            });
            paginationElement.appendChild(prevButton);
        }
        
        // Page buttons - show limited number for better performance
        const maxButtons = 5;
        let startPage = Math.max(1, currentPage - Math.floor(maxButtons / 2));
        let endPage = Math.min(totalPages, startPage + maxButtons - 1);
        
        if (endPage - startPage + 1 < maxButtons && startPage > 1) {
            startPage = Math.max(1, endPage - maxButtons + 1);
        }
        
        for (let i = startPage; i <= endPage; i++) {
            const pageButton = document.createElement('button');
            pageButton.textContent = i;
            pageButton.className = i === currentPage ? 'active' : '';
            pageButton.addEventListener('click', () => displayRecordsPage(i));
            paginationElement.appendChild(pageButton);
        }
        
        // Next button
        if (totalPages > 1) {
            const nextButton = document.createElement('button');
            nextButton.textContent = '→';
            nextButton.addEventListener('click', () => {
                if (currentPage < totalPages) {
                    displayRecordsPage(currentPage + 1);
                }
            });
            paginationElement.appendChild(nextButton);
        }
    }
    
    function displayRecordsPage(page) {
        currentPage = page;
        
        const tableBody = document.querySelector('#problem-records-table tbody');
        tableBody.innerHTML = '';
        
        const start = (page - 1) * recordsPerPage;
        const end = Math.min(start + recordsPerPage, filteredRecords.length);
        
        for (let i = start; i < end; i++) {
            const record = filteredRecords[i];
            const row = document.createElement('tr');
            
            row.innerHTML = `
                <td><a href="/Person2/${record.PersonId}" target="_blank">${record.Name}</a></td>
                <td>${record.Age}</td>
                <td>${record.MissingDataTags}</td>
                <td>${record.LastModified}</td>
            `;
            
            tableBody.appendChild(row);
        }
        
        generatePagination();
        
        const recordCountElement = document.getElementById('records-count');
        if (recordCountElement) {
            recordCountElement.textContent = `Showing ${start + 1}-${end} of ${filteredRecords.length} records`;
        }
    }
    
    function filterProblemRecords() {
        const filterValue = document.getElementById('issue-filter').value;
        
        if (filterValue === 'all') {
            filteredRecords = [...problemRecordsData];
        } else {
            filteredRecords = problemRecordsData.filter(record => {
                switch (filterValue) {
                    case 'missing-email':
                        return record.MissingEmail === 1;
                    case 'missing-phone':
                        return record.MissingPhone === 1;
                    case 'missing-address':
                        return record.MissingAddress === 1;
                    case 'missing-gender':
                        return record.MissingGender === 1;
                    case 'missing-birthdate':
                        return record.MissingBirthDate === 1;
                    default:
                        return true;
                }
            });
        }
        
        currentPage = 1;
        generatePagination();
        displayRecordsPage(1);
    }
    
    function changeRecordsPerPage() {
        recordsPerPage = parseInt(document.getElementById('records-per-page').value);
        currentPage = 1;
        generatePagination();
        displayRecordsPage(1);
    }
    
    // Enhanced table sorting with performance optimizations
    function sortTable(tableId, columnIndex, isNumeric = false) {
        // Basic setup
        const table = document.getElementById(tableId);
        const tbody = table.querySelector('tbody');
        const rows = Array.from(tbody.querySelectorAll('tr'));
        
        // Determine sort direction
        const headerCells = table.querySelectorAll('th');
        const headerCell = headerCells[columnIndex - 1]; // Adjust index to be 0-based
        const currentDir = headerCell.getAttribute('data-sort') || 'asc';
        const newDir = currentDir === 'asc' ? 'desc' : 'asc';
        
        // Update sort indicators on headers
        headerCells.forEach(cell => cell.setAttribute('data-sort', ''));
        headerCell.setAttribute('data-sort', newDir);
        
        // Debug info
        // console.log(`Sorting table ${tableId}, column ${columnIndex}, isNumeric: ${isNumeric}, direction: ${newDir}`);
        
        // Sort the rows
        rows.sort((a, b) => {
            // Get cell content
            const aCell = a.cells[columnIndex - 1];
            const bCell = b.cells[columnIndex - 1];
            
            if (!aCell || !bCell) return 0;
            
            const aContent = aCell.textContent.trim();
            const bContent = bCell.textContent.trim();
            
            // Log the raw values we're comparing
            // console.log(`Comparing raw: "${aContent}" vs "${bContent}"`);
            
            // For numeric columns (like percentages)
            if (isNumeric) {
                // Extract numeric values (handling percentages)
                // Convert "83.6%" to 83.6
                const aValue = parseFloat(aContent.replace('%', ''));
                const bValue = parseFloat(bContent.replace('%', ''));
                
                // Log the extracted numeric values
                // console.log(`Comparing numeric: ${aValue} vs ${bValue}`);
                
                // Compare the numeric values directly
                if (!isNaN(aValue) && !isNaN(bValue)) {
                    return newDir === 'asc' ? aValue - bValue : bValue - aValue;
                }
            }
            
            // Fallback to simple string comparison
            return newDir === 'asc' ? 
                aContent.localeCompare(bContent) : 
                bContent.localeCompare(aContent);
        });
        
        // Clear the table and add the sorted rows
        while (tbody.firstChild) {
            tbody.removeChild(tbody.firstChild);
        }
        
        rows.forEach(row => tbody.appendChild(row));
    }
    
    // Performance optimization for large tables
    function optimizeTable(tableId, maxInitialRows = 100) {
        const table = document.getElementById(tableId);
        if (!table) return;
        
        const tbody = table.querySelector('tbody');
        const allRows = Array.from(tbody.querySelectorAll('tr'));
        
        if (allRows.length <= maxInitialRows) return;
        
        // Hide rows beyond the initial set
        allRows.slice(maxInitialRows).forEach(row => {
            row.style.display = 'none';
        });
        
        // Add a "Load more" button
        const loadMoreRow = document.createElement('tr');
        const loadMoreCell = document.createElement('td');
        loadMoreCell.colSpan = table.querySelector('tr').cells.length;
        loadMoreCell.style.textAlign = 'center';
        
        const loadMoreBtn = document.createElement('button');
        loadMoreBtn.textContent = 'Load more records';
        loadMoreBtn.className = 'action-btn';
        loadMoreBtn.addEventListener('click', () => {
            // Show the next batch of rows
            const hiddenRows = Array.from(tbody.querySelectorAll('tr[style="display: none;"]'));
            const nextBatch = hiddenRows.slice(0, maxInitialRows);
            
            nextBatch.forEach(row => {
                row.style.display = '';
            });
            
            // Remove the "Load more" button if all rows are visible
            if (hiddenRows.length <= maxInitialRows) {
                loadMoreRow.remove();
            }
        });
        
        loadMoreCell.appendChild(loadMoreBtn);
        loadMoreRow.appendChild(loadMoreCell);
        tbody.appendChild(loadMoreRow);
    }
    //Data Score toolTip
    function showDataCompletenessTooltip(event) {
        // Remove any existing tooltips
        const existingTooltips = document.querySelectorAll('.tooltip-popup');
        existingTooltips.forEach(tooltip => tooltip.remove());
        
        // Create tooltip element
        const tooltip = document.createElement('div');
        tooltip.className = 'tooltip-popup';
        tooltip.innerHTML = `
            <h3>Data Completeness Score Explained</h3>
            <p>The Data Completeness Score represents the overall quality of your people records, calculated by tracking the presence of critical information.</p>
            
            <h4>How It's Calculated</h4>
            <ul>
                <li>A percentage that reflects how much critical information is present in a person's record</li>
                <li>Checks key fields: Gender, Birth Date, Address, Phone, Email, Member Status</li>
                <li>Calculates the percentage of <strong>complete</strong> records</li>
                <li>Ranges from 0-100%</li>
            </ul>
            
            <h4>Score Ratings</h4>
            <ul>
                <li>95-100%: <strong style="color:#4caf50;">Excellent</strong></li>
                <li>85-94%: <strong style="color:#8bc34a;">Good</strong></li>
                <li>70-84%: <strong style="color:#ffc107;">Fair</strong></li>
                <li>Below 70%: <strong style="color:#f44336;">Poor</strong></li>
            </ul>
            
            <p><em>Goal: Maintain a high score by ensuring most records have complete, essential information.</em></p>

            <h4>Example:</h4>
            <ul>
                <li>Gender (5% missing)</li>
                <li>Birth Date (10% missing)</li>
                <li>Address (15% missing)</li>
                <li>Phone (20% missing)</li>
                <li>Email (25% missing)</li>
            </ul>
            <p>The average missing percentage would be <br><br>(5+10+15+20+25)/5 = 15%<br><br>
            100% - 15% = 85% ("Good" Data Completeness Score)<br>
            
            <button onclick="this.parentElement.remove();" style="
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                cursor: pointer;
                margin-top: 10px;
            ">Close</button>
        `;
        
        // Position the tooltip near the clicked icon
        tooltip.style.top = `${event.clientY + 10}px`;
        tooltip.style.left = `${event.clientX - 200}px`;
        
        // Add to body and make visible
        document.body.appendChild(tooltip);
        tooltip.style.display = 'block';
        
        // Close tooltip if clicked outside
        function closeTooltipHandler(e) {
            if (!tooltip.contains(e.target)) {
                tooltip.remove();
                document.removeEventListener('click', closeTooltipHandler);
            }
        }
        
        // Add click listener to document to close tooltip
        setTimeout(() => {
            document.addEventListener('click', closeTooltipHandler);
        }, 0);
    }


    // Phone Email pop-up
    function showPhoneEmailTooltip(event, type) {
        // Remove any existing tooltips
        const existingTooltips = document.querySelectorAll('.tooltip-popup');
        existingTooltips.forEach(tooltip => tooltip.remove());
        
        // Create tooltip element
        const tooltip = document.createElement('div');
        tooltip.className = 'tooltip-popup';
        
        // Define tooltip content based on type
        const tooltipContent = type === 'phone' 
            ? `
                <h3>Phone Number Tracking</h3>
                <p>Phone numbers are not tracked for individuals under 13 years old due to privacy considerations.</p>
                
                <h4>What This Means</h4>
                <ul>
                    <li>Children under 13 are excluded from phone number statistics</li>
                    <li>This ensures compliance with child privacy guidelines</li>
                    <li>Parents can add contact information as needed</li>
                </ul>
            `
            : `
                <h3>Email Address Tracking</h3>
                <p>Email addresses are not tracked for individuals under 13 years old due to privacy regulations.</p>
                
                <h4>Key Points</h4>
                <ul>
                    <li>Children under 13 are excluded from email statistics</li>
                    <li>This protects minors' online privacy</li>
                    <li>Parents can manage contact information</li>
                </ul>
            `;
        
        // Set the tooltip content
        tooltip.innerHTML = tooltipContent + `
            <button onclick="this.parentElement.remove();" style="
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                cursor: pointer;
                margin-top: 10px;
            ">Close</button>
        `;
        
        // Position the tooltip near the clicked icon
        tooltip.style.top = `${event.clientY + 10}px`;
        tooltip.style.left = `${event.clientX - 200}px`;
        
        // Add to body and make visible
        document.body.appendChild(tooltip);
        tooltip.style.display = 'block';
        
        // Close tooltip if clicked outside
        function closeTooltipHandler(e) {
            if (!tooltip.contains(e.target)) {
                tooltip.remove();
                document.removeEventListener('click', closeTooltipHandler);
            }
        }
        
        // Add click listener to document to close tooltip
        setTimeout(() => {
            document.addEventListener('click', closeTooltipHandler);
        }, 0);
    }

    // Performance monitoring utility
    function trackPerformance() {
        const metrics = {
            initialLoadTime: 0,
            chartRenderTime: 0,
            tableLoadTime: 0
        };
        
        const startTime = performance.now();
        
        window.addEventListener('load', () => {
            metrics.initialLoadTime = performance.now() - startTime;
            console.log(`Initial page load: ${metrics.initialLoadTime.toFixed(2)}ms`);
        });
        
        // Add observer for tracking visible elements
        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    // When overview tab becomes visible, initialize charts
                    if (entry.target.id === 'overview' && !chartsInitialized) {
                        const chartStart = performance.now();
                        initializeCharts();
                        metrics.chartRenderTime = performance.now() - chartStart;
                        console.log(`Chart rendering: ${metrics.chartRenderTime.toFixed(2)}ms`);
                    }
                    
                    // When records tab becomes visible, load problem records
                    if (entry.target.id === 'records' && !problemRecordsLoaded) {
                        const tableStart = performance.now();
                        loadProblemRecords();
                        // We'll log the table load time when the data is actually loaded
                        setTimeout(() => {
                            metrics.tableLoadTime = performance.now() - tableStart;
                            console.log(`Table data loading: ${metrics.tableLoadTime.toFixed(2)}ms`);
                        }, 500);
                    }
                    
                    observer.unobserve(entry.target);
                }
            });
        }, { threshold: 0.1 });
        
        // Observe the tab content elements
        document.querySelectorAll('.tab-content').forEach(content => {
            observer.observe(content);
        });
    }
    
    // Initialize dashboard
    document.addEventListener('DOMContentLoaded', function() {
        // Only load critical data initially
        // Initialize overview tab charts since it's active by default
        if (document.getElementById('overview').classList.contains('active')) {
            setTimeout(initializeCharts, 10); // Small delay to ensure DOM is ready
        }
        
        // Apply table optimizations
        optimizeTable('status-table');
        optimizeTable('age-table');
        
        // Start performance tracking
        trackPerformance();
    });
    </script>
</body>
</html>
'''

# Execute queries
overall_stats = q.QuerySql(sqlDataQuality)
problem_records = q.QuerySql(sqlMissingDataPeople)

# Process data for display
status_stats = {}
age_stats = {}

for row in overall_stats:
    if row.MetricType == 'OverallStats':
        status_stats[row.RecordStatus] = row
    elif row.MetricType == 'AgeBreakdown':
        age_stats[row.RecordStatus] = row

# Generate status table rows
status_table_rows = ''
status_labels = []
status_data = []

for status in status_stats:
    stats = status_stats[status]
    status_labels.append(status)
    status_data.append(stats.TotalRecords)
    
    # Format data for table row
    score_class = ''
    if stats.DataCompletenessScore < 70:
        score_class = 'danger'
    elif stats.DataCompletenessScore < 85:
        score_class = 'warning'
        
    status_table_rows += '''
    <tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}%</td>
        <td>{3}%</td>
        <td>{4}%</td>
        <td>{5}%</td>
        <td>{6}%</td>
        <td class="{7}">{8:.1f}%</td>
    </tr>
    '''.format(
        status,
        "{:,}".format(stats.TotalRecords),
        stats.PctMissingGender,
        stats.PctMissingBirthDate,
        stats.PctMissingAddress,
        stats.PctMissingPhone,
        stats.PctMissingEmail,
        score_class,
        stats.DataCompletenessScore
    )

# Generate age group table rows
age_table_rows = ''
age_group_labels = []
age_group_scores = []
age_group_counts = []

# Define a sort order for age groups
age_order = {
    'Under 13': 0,
    '13-17': 1,
    '18-24': 2,
    '25-34': 3,
    '35-44': 4,
    '45-54': 5,
    '55-64': 6,
    '65-74': 7,
    '75+': 8,
    'Unknown': 9
}

sorted_age_groups = sorted(age_stats.items(), key=lambda x: age_order.get(x[0], 10))

for age_item in sorted_age_groups:
    age = age_item[0]
    stats = age_item[1]
    
    if age != 'Unknown' or stats.TotalRecords > 0:  # Only include Unknown if it has records
        age_group_labels.append(age)
        age_group_scores.append(round(stats.DataCompletenessScore, 1))
        age_group_counts.append(stats.TotalRecords)
        
        # Format data for table row
        score_class = ''
        if stats.DataCompletenessScore < 70:
            score_class = 'danger'
        elif stats.DataCompletenessScore < 85:
            score_class = 'warning'
            
        age_table_rows += '''
        <tr>
            <td>{0}</td>
            <td>{1}</td>
            <td>{2}%</td>
            <td>{3}%</td>
            <td>{4}%</td>
            <td>{5}%</td>
            <td class="{6}">{7:.1f}%</td>
        </tr>
        '''.format(
            age,
            "{:,}".format(stats.TotalRecords),
            stats.PctMissingGender,
            stats.PctMissingAddress,
            stats.PctMissingPhone,
            stats.PctMissingEmail,
            score_class,
            stats.DataCompletenessScore
        )

# Problem records table generation section with fix
problem_records_rows = ''


for record in problem_records:
    # Create missing data tags
    missing_data_tags = []

    # Add missing data tags for ALL types of missing data
    if record.MissingGender == 1:
        missing_data_tags.append('<span class="badge badge-warning">Gender</span>')
    if record.MissingBirthDate == 1:
        missing_data_tags.append('<span class="badge badge-warning">Birth Date</span>')
    if record.MissingMaritalStatus == 1:
        missing_data_tags.append('<span class="badge badge-info">Marital Status</span>')
    if record.MissingAddress == 1:
        missing_data_tags.append('<span class="badge badge-danger">Address</span>')
    if record.MissingPhone == 1:
        missing_data_tags.append('<span class="badge badge-danger">Phone</span>')
    if record.MissingEmail == 1:
        missing_data_tags.append('<span class="badge badge-danger">Email</span>')
    if record.MissingPhoto == 1:
        missing_data_tags.append('<span class="badge badge-info">Photo</span>')
    
    # Format the LastModified date safely
    last_modified_str = str(record.LastModified) if record.LastModified else 'N/A'
    
    # Always add the record, even if no missing data tags
    problem_records_rows += '''
    <tr>
        <td><a href="/Person2/{0}" target="_blank">{1}</a></td>
        <td>{2}</td>
        <td>{3}</td>
        <td>{4}</td>
    </tr>
    '''.format(
        record.PeopleId,
        record.Name,
        record.Age,
        ''.join(missing_data_tags) if missing_data_tags else 'No missing data detected',
        last_modified_str
    )


# Calculate missing data percentages for active records
active_stats = status_stats.get('Active', {})
missing_data_percentages = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]



# Calculate missing data percentages for active records
if active_stats:
    missing_data_percentages = [
        round(float(active_stats.PctMissingGender or 0.0), 2),
        round(float(active_stats.PctMissingBirthDate or 0.0), 2),
        round(float(active_stats.PctMissingAddress or 0.0), 2),
        round(float(active_stats.PctMissingPhone or 0.0), 2),
        round(float(active_stats.PctMissingEmail or 0.0), 2),
        round(float(active_stats.PctMissingPhoto or 0.0), 2),
        #round(float(active_stats.MissingMemberStatus * 100.0 / active_stats.TotalRecords or 0.0), 2)
    ]
else:
    missing_data_percentages = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

# Ensure all values are floats and round to 2 decimal places
missing_data_percentages = [
    round(float(x), 2) if x is not None and x is not False else 0.0 
    for x in missing_data_percentages
]

    #print missing_data_percentages

# Prepare age group counts (add this section)
age_group_counts = [
    stats.TotalRecords 
    for age, stats in sorted_age_groups 
    if age != 'Unknown' or stats.TotalRecords > 0
]

# Prepare contact information stats
contact_data = []
if active_stats and active_stats.TotalRecords > 0:
    #total_records = active_stats.TotalRecords
    total_records = sum(stats.TotalRecords for status, stats in status_stats.items() if status == 'Active')
    
    # Estimate contact information categories
    #missing_email = active_stats.MissingEmail
    #missing_phone = active_stats.MissingPhone
    #missing_address = active_stats.MissingAddress
    missing_email = status_stats['Active'].MissingEmail
    missing_phone = status_stats['Active'].MissingPhone
    missing_address = status_stats['Active'].MissingAddress
    
    # Estimate overlap (missing both email and phone)
    missing_both = min(missing_email, missing_phone) // 2
    
    # Calculate categories
    has_all = max(0, total_records - (missing_email + missing_phone + missing_address - missing_both))
    missing_email_only = max(0, missing_email - missing_both)
    missing_phone_only = max(0, missing_phone - missing_both)
    missing_address_only = max(0, missing_address)
    multiple_missing = max(0, missing_both)
    
    contact_data = [
        has_all,                  # Has all contact info
        missing_email_only,       # Missing just email
        missing_phone_only,       # Missing just phone
        missing_address_only,     # Missing address
        multiple_missing          # Missing multiple types of contact info
    ]
else:
    contact_data = [0, 0, 0, 0, 0]

# Generate recommended actions
action_items = ''

# Calculate the overall data quality score
active_data_score = 0
if active_stats:
    active_data_score = active_stats.DataCompletenessScore

if active_data_score < 70:
    health_rating = "Poor"
elif active_data_score < 85:
    health_rating = "Fair"
elif active_data_score < 95:
    health_rating = "Good"
else:
    health_rating = "Excellent"

# Count critical actions
critical_count = 0
cleanup_count = 0

# Add action items based on data quality issues
if active_stats:
    # Missing contact information
    if active_stats.MissingEmail > 0:
        critical_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-danger">High</span></td>
            <td>Collect missing email addresses</td>
            <td>Improves digital communication</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-MissingEmailList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.MissingEmail))
        
    if active_stats.MissingPhone > 0:
        critical_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-danger">High</span></td>
            <td>Collect missing phone numbers</td>
            <td>Enables voice/SMS contact</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-MissingPhoneList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.MissingPhone))
        
    if active_stats.MissingAddress > 0:
        critical_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-danger">High</span></td>
            <td>Collect missing addresses</td>
            <td>Enables mail communication</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-MissingAddressList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.MissingAddress))

    # Missing demographic information
    if active_stats.MissingGender > 0:
        cleanup_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-warning">Medium</span></td>
            <td>Update missing gender information</td>
            <td>Improves demographic analysis</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-MissingGenderList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.MissingGender))
        
    if active_stats.MissingBirthDate > 0:
        cleanup_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-warning">Medium</span></td>
            <td>Update missing birth dates</td>
            <td>Enables age-based ministry</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-MissingBirthDateList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.MissingBirthDate))
        
    if active_stats.MissingPhoto > 0:
        cleanup_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-info">Low</span></td>
            <td>Add missing profile photos</td>
            <td>Improves member directory</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-MissingPhotoList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.MissingPhoto))

    if active_stats.BadAddress > 0:
        cleanup_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-warning">Medium</span></td>
            <td>Fix bad addresses</td>
            <td>Reduces returned mail</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-BadAddressList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.BadAddress))
        
    if active_stats.StaleRecord > 0:
        cleanup_count += 1
        action_items += '''
        <tr>
            <td><span class="badge badge-info">Low</span></td>
            <td>Verify stale records</td>
            <td>Updates outdated information</td>
            <td>{0}</td>
            <td><a href="javascript:void(0)" onclick="window.open('/RunScript/TPx_DQD-StaleRecordList', '_blank', 'width=800,height=600')" class="action-btn">Create List</a></td>
        </tr>
        '''.format("{:,}".format(active_stats.StaleRecord))

# Prepare data for dashboard
active_count = 0
if 'Active' in status_stats:
    active_count = status_stats['Active'].TotalRecords

data_completeness_score = "0.0"
if active_stats:
    data_completeness_score = "{:.1f}".format(active_stats.DataCompletenessScore)

# Determine score rating
score_rating = '<span style="color: #f44336;">Poor</span>'
if active_stats:
    if active_stats.DataCompletenessScore >= 95:
        score_rating = '<span style="color: #4caf50;">Excellent</span>'
    elif active_stats.DataCompletenessScore >= 85:
        score_rating = '<span style="color: #8bc34a;">Good</span>'
    elif active_stats.DataCompletenessScore >= 70:
        score_rating = '<span style="color: #ffc107;">Fair</span>'

# Calculate missing critical data count
missing_critical_count = 0
if active_stats:
    # Count people missing at least one of: address, phone, email
    missing_critical_count = active_stats.MissingAddress + active_stats.MissingPhone + active_stats.MissingEmail
    # Adjust for potential overlap (rough estimate)
    missing_critical_count = min(missing_critical_count, active_stats.TotalRecords)

# Count stale records
stale_records = 0
if active_stats:
    stale_records = active_stats.StaleRecord

# Format all chart data for JavaScript
status_labels_str = str(status_labels)
status_data_str = str([stats.TotalRecords for stats in status_stats.values()])


def safe_json_convert(data):
    """
    Recursively convert data to a JSON-serializable format
    """
    if isinstance(data, (int, long, str, unicode)):
        return data
    elif isinstance(data, float):
        return round(data, 2)
    elif isinstance(data, list):
        return [safe_json_convert(item) for item in data]
    elif isinstance(data, dict):
        return {key: safe_json_convert(value) for key, value in data.items()}
    else:
        return str(data)

# Modify how you prepare the missing_data_percentages
missing_data_percentages = [
    round(float(x), 2) if x is not None else 0.0 
    for x in missing_data_percentages
]

# Replace template variables with safe JSON serialization
template_variables = {
    'active_count': "{:,}".format(active_count),
    'data_completeness_score': data_completeness_score,
    'score_rating': score_rating,
    'missing_critical_count': "{:,}".format(missing_critical_count),
    'stale_records': "{:,}".format(stale_records),
    'status_table_rows': status_table_rows,
    'age_table_rows': age_table_rows,
    'problem_records_rows': problem_records_rows,
    'health_rating': health_rating,
    'critical_actions': str(critical_count),
    'cleanup_opportunities': str(cleanup_count),
    'action_items': action_items,
    'status_labels': safe_json_convert(status_labels),
    'status_data': safe_json_convert(status_data),
    'missing_data_percentages': safe_json_convert(missing_data_percentages),
    'age_group_labels': safe_json_convert(age_group_labels),
    'age_group_scores': safe_json_convert(age_group_scores),
    'contact_data': safe_json_convert(contact_data),
    'age_group_counts': safe_json_convert(age_group_counts),
    'contact_data': safe_json_convert(contact_data)
}

# Replace template variables
dashboard_html = html_template
for key, value in template_variables.items():
    # Convert to JSON string for most variables
    if key in ['status_labels', 'status_data', 'missing_data_percentages', 
               'age_group_labels', 'age_group_scores', 'contact_data']:
        value = json.dumps(value)
    
    dashboard_html = dashboard_html.replace('{{' + key + '}}', str(value))

print(dashboard_html)
