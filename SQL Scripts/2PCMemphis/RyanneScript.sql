-- Define variables
DECLARE @qtagid int = 347; -- Replace with your actual value
DECLARE @GroupPrefix VARCHAR(50) = 'SG: %';

-- temp table for core data 
IF OBJECT_ID('tempdb..#CoreData') IS NOT NULL DROP TABLE #CoreData;
CREATE TABLE #CoreData (
    PeopleId INT PRIMARY KEY,
    FamilyId INT,
    MemberStatusId INT,
    PositionInFamilyId INT,
    LastName VARCHAR(100),
    PreferredName VARCHAR(100),
    Age INT,
    EmailAddress VARCHAR(255),
    EmailAddress2 VARCHAR(255),
    SendEmailAddress1 BIT,
    SendEmailAddress2 BIT,
    HomePhone VARCHAR(50),
    CellPhone VARCHAR(50),
    DeceasedDate DATETIME
);

-- core data table with people in the tag
INSERT INTO #CoreData
SELECT 
    p.PeopleId, p.FamilyId, p.MemberStatusId, p.PositionInFamilyId,
    p.LastName, p.PreferredName, p.Age, p.EmailAddress, p.EmailAddress2,
    p.SendEmailAddress1, p.SendEmailAddress2, p.HomePhone, p.CellPhone, p.DeceasedDate
FROM People p
JOIN TagPerson tp ON tp.PeopleId = p.PeopleId
WHERE tp.Id = @qtagid;

-- family members of tagged people
INSERT INTO #CoreData
SELECT 
    p.PeopleId, p.FamilyId, p.MemberStatusId, p.PositionInFamilyId,
    p.LastName, p.PreferredName, p.Age, p.EmailAddress, p.EmailAddress2,
    p.SendEmailAddress1, p.SendEmailAddress2, p.HomePhone, p.CellPhone, p.DeceasedDate
FROM People p
JOIN (SELECT DISTINCT FamilyId FROM #CoreData) f ON p.FamilyId = f.FamilyId
WHERE NOT EXISTS (SELECT 1 FROM #CoreData WHERE PeopleId = p.PeopleId);

-- family data
IF OBJECT_ID('tempdb..#FamilyData') IS NOT NULL DROP TABLE #FamilyData;
CREATE TABLE #FamilyData (
    FamilyId INT PRIMARY KEY,
    HeadOfHouseholdId INT,
    HeadOfHouseholdSpouseId INT,
    AddressLineOne VARCHAR(255),
    AddressLineTwo VARCHAR(255),
    CityName VARCHAR(100),
    StateCode VARCHAR(20),
    ZipCode VARCHAR(20)
);

INSERT INTO #FamilyData
SELECT 
    FamilyId, HeadOfHouseholdId, HeadOfHouseholdSpouseId,
    AddressLineOne, AddressLineTwo, CityName, StateCode, ZipCode
FROM Families
WHERE FamilyId IN (SELECT DISTINCT FamilyId FROM #CoreData);

-- original CTEs but with joins to the temp table
WITH 
CC AS (
    SELECT 
        O.OrganizationName,
        O.OrganizationId,
        o.LeaderName,
        P.PeopleId
    FROM #CoreData P
    CROSS APPLY (
        SELECT TOP 1 
            O.OrganizationName,
            O.OrganizationID,
            O.LeaderName
        FROM OrganizationMembers OM
        JOIN Organizations O ON O.OrganizationId = OM.OrganizationId
        JOIN DivOrg DO ON Do.OrgId = O.OrganizationId
        WHERE OM.PeopleId = P.PeopleId AND (do.DivId = 23)  
        ORDER BY O.OrganizationId
    ) O
),
Parish AS (
    SELECT 
        O.OrganizationName AS ParishName,
        o.IsLeader,
        P.PeopleID,
        p.FamilyID,
        o.GroupName
    FROM #CoreData P
    CROSS APPLY (
        SELECT TOP 1 
            O.OrganizationName,
            CASE WHEN mt.AttendanceTypeId = 10 THEN 'Leader' ELSE '' END as IsLeader,
            D.Name,
            mt1.Name as GroupName
        FROM OrganizationMembers OM
        JOIN Organizations O ON O.OrganizationId = OM.OrganizationId
        JOIN DivOrg do ON do.orgID = o.organizationID
        JOIN Division D ON D.Id = do.DivId
        LEFT JOIN OrgMemMemTags ommt ON ommt.OrgId = om.OrganizationId 
            AND ommt.PeopleId = om.PeopleId 
            AND ommt.MemberTagId IN (SELECT ID FROM MemberTags WHERE Name LIKE @GroupPrefix)
        LEFT JOIN MemberTags mt1 ON mt1.Id = ommt.MemberTagId and mt1.Name LIKE @GroupPrefix
        JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
        WHERE do.DivId IN (126,127,128,129) AND OM.PeopleId = P.PeopleId
        ORDER BY O.OrganizationId
    ) O
),
Officers AS (
    SELECT DISTINCT
        OM.PeopleId,
        MIN(O.OrganizationId) AS OrganizationId,
        MIN(O.OrganizationName) AS OfficerType
    FROM ProgDiv PD 
    JOIN DivOrg DO ON DO.DivId = PD.DivId AND PD.ProgId = 1112 AND DO.OrgId IN (42,177,178,201,218,298,299)
    JOIN Organizations O ON O.OrganizationId = Do.OrgId 
    JOIN OrganizationMembers OM ON OM.OrganizationId = O.OrganizationId
    JOIN #CoreData P on OM.PeopleId = P.PeopleId
    WHERE P.MemberStatusId <> 110 AND OM.MemberTypeID = 220
    GROUP BY OM.PeopleId
),
AddressInfo AS (
    SELECT 
        f.FamilyId,
        TRIM('.' FROM REPLACE(ai.RegionName, ' Extd', '')) AS RegionName
    FROM #FamilyData f
    LEFT JOIN AddressInfo ai ON ai.FamilyId = f.FamilyId
),
FamilyGroups AS (
    SELECT 
        f.FamilyId,
        STRING_AGG(
            CASE 
                WHEN p.PositionInFamilyId NOT IN (10, 20) AND p.DeceasedDate IS NULL 
                THEN p.PreferredName + CASE WHEN p.MemberStatusId = 10 THEN '*' ELSE '' END 
            END, 
            ', '
        ) WITHIN GROUP (ORDER BY p.PositionInFamilyId) AS Kids,
        
        STRING_AGG(
            CASE 
                WHEN p.PositionInFamilyId = 10 AND p.DeceasedDate IS NULL
                THEN p.PreferredName + CASE WHEN p.MemberStatusId = 10 THEN '*' ELSE '' END 
            END, 
            ', '
        ) WITHIN GROUP (ORDER BY p.PositionInFamilyId) AS PrimaryAdults,
        
        STRING_AGG(
            CASE 
                WHEN p.PositionInFamilyId = 20 AND p.DeceasedDate IS NULL
                THEN p.PreferredName + ' ' + p.LastName + CASE WHEN p.MemberStatusId = 10 THEN '*' ELSE '' END 
            END, 
            ', '
        ) WITHIN GROUP (ORDER BY p.PositionInFamilyId) AS Other,
        
        STRING_AGG(
            CASE WHEN p.DeceasedDate IS NULL THEN p.PreferredName END, 
            ', '
        ) WITHIN GROUP (ORDER BY p.PositionInFamilyId) AS Firstnames
    FROM #CoreData p
    JOIN #FamilyData f ON f.FamilyId = p.FamilyId
    GROUP BY f.FamilyId
),
Adults AS (
    SELECT DISTINCT
        ISNULL(SUBSTRING(
            (
                SELECT DISTINCT ', ' + pa.ParishName AS [text()]
                FROM Parish Pa
                JOIN #CoreData pe ON pe.PeopleId = pa.PeopleId
                WHERE Pa.FamilyId = p.FamilyId AND pe.PositionInFamilyId = 10
                FOR XML PATH ('')
            ), 2, 1000), 'None') AS ParishName1,
        ISNULL(ISNULL(hpa.ParishName, ISNULL(spa.ParishName, ppa.ParishName)), 'None') AS ParishName,
        ISNULL(ISNULL(hpa.GroupName, ISNULL(spa.GroupName, ppa.GroupName)), 'SG: unassigned') AS GroupName,
        ISNULL(ai.RegionName, 'None') AS LivesIn,
        ISNULL(hpa.IsLeader, '') AS IsLeader,
        F.FamilyId,
        h.PeopleId AS MainID,
        ISNULL(CAST(S.PeopleId AS VARCHAR(10)), '') AS SPSID,
        h.PreferredName + ISNULL(' & ' + s.PreferredName, '') AS PreferredNames,
        H.LastName,
        CASE 
            WHEN H.MemberStatusId = 110 THEN 'Pastor'
            WHEN OH.OfficerType LIKE '%Elder%' AND OH.OfficerType LIKE '%Deacon%' THEN 'Elder'
            WHEN OS.OfficerType LIKE '%Deacon%' AND OH.OfficerType LIKE '%Elder%' THEN 'Elder'
            WHEN OH.OfficerType LIKE '%Elder%' THEN 'Elder'
            WHEN OH.OfficerType LIKE '%Deacon%' THEN 'Deacon'
            WHEN OS.OfficerType LIKE '%Deacon%' THEN 'Deacon'
            WHEN h.MemberStatusId = 10 THEN 'Member'
            WHEN s.MemberStatusId = 10 THEN 'Member'
            WHEN p.MemberStatusId = 10 AND p.PositionInFamilyId = 30 THEN 'Member'
            WHEN p.MemberStatusId = 10 AND p.PositionInFamilyId = 20 THEN 'Member'
            ELSE 'Non-Member'
        END AS OfficerType,
        CASE 
            WHEN h.SendEmailAddress1 = 1 AND ISNULL(h.EmailAddress, '') <> '' THEN h.EmailAddress
            WHEN h.SendEmailAddress1 = 0 AND h.SendEmailAddress2 = 1 AND ISNULL(h.EmailAddress2, '') <> '' THEN h.EmailAddress2 
            ELSE ''
        END AS HoHEmail,
        CASE 
            WHEN S.SendEmailAddress1 = 1 AND ISNULL(s.EmailAddress, '') <> '' THEN S.EmailAddress
            WHEN s.SendEmailAddress1 = 0 AND s.SendEmailAddress2 = 1 AND ISNULL(s.EmailAddress2, '') <> '' THEN s.EmailAddress2 
            ELSE ''
        END AS SpouseEmail,
        CASE 
            WHEN LEN(h.homephone) = 10 THEN '(' + SUBSTRING(h.HomePhone, 1, 3) + ') ' + SUBSTRING(h.homephone, 4, 3) + '-' + SUBSTRING(h.homephone, 7, 4) 
            WHEN LEN(h.homephone) = 7 THEN LEFT(h.homephone, 3) + '-' + RIGHT(h.homephone, 4)
            ELSE ''
        END AS HomePhone,
        CASE 
            WHEN LEN(h.CellPhone) = 10 THEN '(' + SUBSTRING(h.CellPhone, 1, 3) + ') ' + SUBSTRING(h.CellPhone, 4, 3) + '-' + SUBSTRING(h.CellPhone, 7, 4) 
            WHEN LEN(h.CellPhone) = 7 THEN LEFT(h.CellPhone, 3) + '-' + RIGHT(h.CellPhone, 4)
            ELSE ''
        END AS HOHCell,
        CASE 
            WHEN LEN(s.CellPhone) = 10 THEN '(' + SUBSTRING(s.CellPhone, 1, 3) + ') ' + SUBSTRING(s.CellPhone, 4, 3) + '-' + SUBSTRING(s.CellPhone, 7, 4) 
            WHEN LEN(s.CellPhone) = 7 THEN LEFT(s.CellPhone, 3) + '-' + RIGHT(s.CellPhone, 4)
            ELSE ''
        END AS SpouseCell,
        F.AddressLineOne AS address1,
        F.AddressLineTwo AS Address2,
        F.CityName AS City,
        F.StateCode AS [State],
        CASE 
            WHEN LEN(F.ZipCode) = 9 THEN LEFT(F.ZipCode, 5) + '-' + RIGHT(F.ZipCode, 4)
            ELSE F.ZipCode
        END AS ZipCode,
        CASE 
            WHEN CCH.OrganizationName IS NULL THEN ISNULL(CCS.OrganizationName, '')
            ELSE ISNULL(CCH.OrganizationName, '')
        END AS CC,
        CASE 
            WHEN CCH.LeaderName IS NULL THEN ISNULL(CCS.LeaderName, '')
            ELSE CCH.LeaderName
        END AS CCLeader,
        H.PreferredName AS HOHFirstName,
        ISNULL(S.PreferredName, '') AS SpouseFirstName,
        h.lastname + ', ' + h.PreferredName + ISNULL(' & ' + s.PreferredName, '') AS Names1,
        h.PreferredName + ISNULL(' & ' + s.PreferredName, '') AS Names2,
        'https://my.2pc.org/Person2/' + CONVERT(VARCHAR(50), f.HeadOfHouseholdID) AS Link,
        CASE 
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) > 0 AND ISNULL(H.Age, ISNULL(S.Age, 0)) < 18 THEN 'Under 18'
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) >= 18 AND ISNULL(H.Age, ISNULL(S.Age, 0)) < 25 THEN '18-24'
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) >= 25 AND ISNULL(H.Age, ISNULL(S.Age, 0)) < 35 THEN '25-34'
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) >= 35 AND ISNULL(H.Age, ISNULL(S.Age, 0)) < 45 THEN '35-44'
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) >= 45 AND ISNULL(H.Age, ISNULL(S.Age, 0)) < 55 THEN '45-54'
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) >= 55 AND ISNULL(H.Age, ISNULL(S.Age, 0)) < 65 THEN '55-64'
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) >= 65 AND ISNULL(H.Age, ISNULL(S.Age, 0)) < 75 THEN '65-74'
            WHEN ISNULL(H.Age, ISNULL(S.Age, 0)) >= 75 THEN '75+'
            ELSE ''
        END AS AgeRange,
        CASE     
            WHEN OH.OfficerType LIKE 'Active%' THEN 'Active'
            WHEN OH.OfficerType LIKE 'Inactive%' THEN 'Inactive'
            WHEN OH.OfficerType LIKE 'Rotated%' THEN 'Rotated'
            WHEN OH.OfficerType LIKE '%Emeriti%' THEN 'Emeriti'
            WHEN OS.OfficerType LIKE 'Active%' THEN 'Active'
            WHEN OS.OfficerType LIKE 'Inactive%' THEN 'Inactive'
            WHEN OS.OfficerType LIKE 'Rotated%' THEN 'Rotated'
            WHEN OS.OfficerType LIKE '%Emeriti%' THEN 'Emeriti'
            ELSE ''
        END AS OfficerStatus,
        CASE     
            WHEN H.MemberStatusId = 110 THEN 'Pastor'
            WHEN OH.OrganizationId IS NOT NULL THEN 'Officer'
            WHEN OS.OrganizationId IS NOT NULL THEN 'Officer'
            ELSE 'Member'
        END AS MemberType,
        fg.Kids,
        fg.PrimaryAdults,
        fg.Other,
        fg.Firstnames,
        CASE 
            WHEN hpa.IsLeader = 'Leader' THEN 1 
            WHEN h.MemberStatusId = 110 THEN 2
            WHEN oh.OfficerType LIKE '%Elder%' THEN 3
            WHEN Oh.OfficerType LIKE '%Deacon%' THEN 4
            WHEN os.OfficerType LIKE '%Deacon%' AND oh.OfficerType NOT LIKE '%Elder%' THEN 4
            ELSE 5
        END AS sort
    FROM #CoreData p
        JOIN #FamilyData F ON F.FamilyId = P.FamilyId
        JOIN #CoreData H ON H.PeopleId = F.HeadOfHouseholdId
        LEFT JOIN #CoreData S ON S.PeopleId = F.HeadOfHouseholdSpouseId
        LEFT JOIN CC CCH ON CCH.PeopleId = H.PeopleId
        LEFT JOIN CC CCS ON CCS.PeopleId = S.PeopleId
        LEFT JOIN Officers OH ON OH.PeopleId = H.PeopleId
        LEFT JOIN Officers OS ON OS.PeopleId = S.PeopleId
        LEFT JOIN Parish hpa ON hpa.PeopleId = h.PeopleId
        LEFT JOIN Parish spa ON spa.PeopleId = s.PeopleId
        LEFT JOIN Parish ppa ON ppa.PeopleId = p.PeopleId
        LEFT JOIN AddressInfo ai ON ai.FamilyId = p.FamilyId
        LEFT JOIN FamilyGroups fg ON fg.FamilyId = p.FamilyId
),
kidsAndOther AS (
    SELECT DISTINCT
        ISNULL(SUBSTRING(
            (
                SELECT ', ' + Pa.ParishName AS [text()]
                FROM Parish Pa
                WHERE Pa.PeopleId = P.PeopleId 
                ORDER BY Pa.ParishName
                FOR XML PATH ('')
            ), 2, 1000), 'None') AS ParishName1,
        ISNULL(ppa.ParishName, 'None') AS ParishName,
        ISNULL(ppa.GroupName, 'SG: unassigned') AS GroupName,
        ISNULL(ai.RegionName, 'None') AS LivesIn,
        ISNULL(ppa.IsLeader, '') AS IsLeader,
        P.FamilyId,
        P.PeopleId AS MainID,
        '' AS SPSID,
        p.PreferredName AS PreferredNames,
        p.LastName,
        CASE 
            WHEN P.MemberStatusId = 110 THEN 'Pastor'
            WHEN OP.OfficerType LIKE '%Elder%' AND OP.OfficerType LIKE '%Deacon%' THEN 'Elder'
            WHEN OP.OfficerType LIKE '%Elder%' THEN 'Elder'
            WHEN OP.OfficerType LIKE '%Deacon%' THEN 'Deacon'
            WHEN p.MemberStatusId = 10 AND p.PositionInFamilyId = 30 THEN 'Member'
            WHEN p.MemberStatusId = 10 AND p.PositionInFamilyId = 20 THEN 'Member'
            ELSE 'Non-Member'
        END AS OfficerType,
        CASE 
            WHEN P.SendEmailAddress1 = 1 AND ISNULL(P.EmailAddress, '') <> '' THEN P.EmailAddress
            WHEN P.SendEmailAddress1 = 0 AND P.SendEmailAddress2 = 1 AND ISNULL(p.EmailAddress2, '') <> '' THEN p.EmailAddress2 
            ELSE ''
        END AS HoHEmail,
        '' AS SpouseEmail,
        CASE 
            WHEN LEN(p.homephone) = 10 THEN '(' + SUBSTRING(p.HomePhone, 1, 3) + ') ' + SUBSTRING(p.homephone, 4, 3) + '-' + SUBSTRING(p.homephone, 7, 4) 
            WHEN LEN(p.homephone) = 7 THEN LEFT(P.homephone, 3) + '-' + RIGHT(P.homephone, 4)
            ELSE ''
        END AS HomePhone,
        CASE 
            WHEN LEN(p.CellPhone) = 10 THEN '(' + SUBSTRING(p.CellPhone, 1, 3) + ') ' + SUBSTRING(p.CellPhone, 4, 3) + '-' + SUBSTRING(p.CellPhone, 7, 4) 
            WHEN LEN(P.CellPhone) = 7 THEN LEFT(p.CellPhone, 3) + '-' + RIGHT(P.CellPhone, 4)
            ELSE ''
        END AS HOHCell,
        '' AS SpouseCell,
        F.AddressLineOne AS address1,
        F.AddressLineTwo AS Address2,
        F.CityName AS City,
        F.StateCode AS [State],
        CASE 
            WHEN LEN(F.ZipCode) = 9 THEN LEFT(F.ZipCode, 5) + '-' + RIGHT(F.ZipCode, 4)
            ELSE F.ZipCode
        END AS ZipCode,
        ISNULL(CCP.OrganizationName, '') AS CC,
        ISNULL(CCP.LeaderName, '') AS CCLeader,
        P.PreferredName AS HOHFirstName,
        '' AS SpouseFirstName,
        p.lastname + ', ' + P.PreferredName AS Names1,
        P.PreferredName AS Names2,
        'https://my.2pc.org/Person2/' + CONVERT(VARCHAR(50), p.PeopleId) AS Link,
        CASE 
            WHEN ISNULL(P.Age, 0) > 0 AND ISNULL(P.Age, 0) < 18 THEN 'Under 18'
            WHEN ISNULL(P.Age, 0) >= 18 AND ISNULL(P.Age, 0) < 25 THEN '18-24'
            WHEN ISNULL(P.Age, 0) >= 25 AND ISNULL(P.Age, 0) < 35 THEN '25-34'
            WHEN ISNULL(P.Age, 0) >= 35 AND ISNULL(P.Age, 0) < 45 THEN '35-44'
            WHEN ISNULL(P.Age, 0) >= 45 AND ISNULL(P.Age, 0) < 55 THEN '45-54'
            WHEN ISNULL(P.Age, 0) >= 55 AND ISNULL(P.Age, 0) < 65 THEN '55-64'
            WHEN ISNULL(P.Age, 0) >= 65 AND ISNULL(P.Age, 0) < 75 THEN '65-74'
            WHEN ISNULL(P.Age, 0) >= 75 THEN '75+'
            ELSE ''
        END AS AgeRange,
        CASE     
            WHEN OP.OfficerType LIKE 'Active%' THEN 'Active'
            WHEN Op.OfficerType LIKE 'Inactive%' THEN 'Inactive'
            WHEN Op.OfficerType LIKE 'Rotated%' THEN 'Rotated'
            WHEN Op.OfficerType LIKE '%Emeriti%' THEN 'Emeriti'
            ELSE ''
        END AS OfficerStatus,
        CASE     
            WHEN p.MemberStatusId = 110 THEN 'Pastor'
            WHEN OP.OrganizationId IS NOT NULL THEN 'Officer'
            WHEN OP.OrganizationId IS NULL AND p.MemberStatusId NOT IN (10, 110) THEN 'Non-Member'
            ELSE 'Member'
        END AS MemberType,
        fg.Kids,
        fg.PrimaryAdults,
        fg.Other,
        fg.Firstnames,
        CASE 
            WHEN ppa.IsLeader = 'Leader' THEN 1 
            WHEN P.MemberStatusId = 110 THEN 2
            WHEN OP.OfficerType LIKE '%Elder%' THEN 3
            WHEN OP.OfficerType LIKE '%Deacon%' THEN 4
            ELSE 5
        END AS sort
    FROM #CoreData P
        JOIN #FamilyData F ON F.FamilyId = P.FamilyId
        LEFT JOIN CC CCP ON CCP.PeopleId = P.PeopleId
        LEFT JOIN #CoreData H ON H.PeopleId = F.HeadOfHouseholdId
        LEFT JOIN #CoreData S ON S.PeopleId = F.HeadOfHouseholdSpouseId
        LEFT JOIN Officers OP ON OP.PeopleId = P.PeopleId
        LEFT JOIN Parish ppa ON ppa.PeopleId = P.PeopleId
        LEFT JOIN AddressInfo ai ON ai.FamilyId = p.FamilyId
        LEFT JOIN FamilyGroups fg ON fg.FamilyId = p.FamilyId
    WHERE 
        p.PositionInFamilyId <> 10 AND (
            (h.MemberStatusId NOT IN (10, 110) AND ISNULL(s.MemberStatusId, 0) NOT IN (10, 110) AND p.MemberStatusId = 10)
        )
),
allpeeps AS (
    SELECT * FROM Adults a JOIN TagPerson tp ON tp.PeopleId = a.MainID AND tp.Id = @qtagid 
    UNION
    SELECT * FROM kidsAndOther k JOIN TagPerson tp ON tp.PeopleId = k.MainID AND tp.Id = @qtagid
    UNION
    SELECT * FROM Adults a JOIN TagPerson tp ON tp.PeopleId = a.SPSID AND tp.Id = @qtagid
),
stuff AS (
    SELECT DISTINCT
        mainID,
        sort,
        ParishName,
        ISNULL(GroupName, 'SG: Unassigned') AS GroupName,
        CONCAT(PreferredNames, ' ', LastName) AS Name,
        LastName,
        PreferredNames,
        OfficerType + CASE WHEN ISNULL(IsLeader, '') = '' THEN '' ELSE ' - ' + IsLeader END AS Role,
        OfficerStatus,
        AgeRange,
        HoHEmail AS [Person/HOH Email],
        SpouseEmail,
        HomePhone,
        HOHCell AS [Person/HOH Cell],
        SpouseCell,
        ISNULL(address1, '') AS addressLine1,
        ISNULL(Address2, '') AS AddressLine2,
        ISNULL(City, '') AS City,
        ISNULL(State, '') AS State,
        ISNULL(ZipCode, '') AS ZipCode,
        CC,
        CCLeader,
        PrimaryAdults AS [Primary Adults  (* indicates Member)],
        Kids AS [Kids  (* indicates Member)],
        Other AS [Other Adults (* indicates Member)]
    FROM allpeeps p
)

SELECT
    mainid,
    ParishName,
    RIGHT(s.GroupName, LEN(s.GroupName) - 4) AS GroupName,
    OfficerStatus,
    LastName,
    PreferredNames,
    Name,
    Role,
    AgeRange,
    [Person/HOH Email],
    SpouseEmail,
    HomePhone,
    [Person/HOH Cell],
    SpouseCell,
    AddressLine1,
    AddressLine2,
    City,
    State,
    ZipCode,
    CC,
    CCLeader,
    [Primary Adults  (* indicates Member)],
    [Kids  (* indicates Member)],
    [Other Adults (* indicates Member)]
FROM stuff s
ORDER BY
    CASE WHEN ParishName = 'None' THEN 'zzz' ELSE ParishName END,  
    GroupName, 
    sort, 
    LastName, 
    PreferredNames;

-- Clean up temp tables
DROP TABLE IF EXISTS #CoreData;
DROP TABLE IF EXISTS #FamilyData;
