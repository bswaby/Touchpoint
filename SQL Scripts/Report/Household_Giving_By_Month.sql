/*==============================================================
  HOUSEHOLD GIVING BY MONTH
  One row per family, one column per month in the range.
==============================================================*/
DECLARE @StartDate DATE = '2025-09-01';
DECLARE @EndDate   DATE = '2026-08-31';
DECLARE @ExcludeFunds VARCHAR(500) = '';

SELECT
    f.FamilyId,
    ISNULL(hoh.Name2, MIN(p.Name2)) AS Household,
    CONVERT(VARCHAR(7), c.ContributionDate, 120) AS Month,
    SUM(c.ContributionAmount) AS Given
FROM dbo.Contribution c WITH (NOLOCK)
JOIN dbo.People   p ON p.PeopleId = c.PeopleId
JOIN dbo.Families f ON f.FamilyId = p.FamilyId
LEFT JOIN dbo.People hoh ON hoh.PeopleId = f.HeadOfHouseholdId
WHERE c.ContributionDate >= @StartDate
  AND c.ContributionDate <  DATEADD(DAY, 1, @EndDate)
  AND c.ContributionStatusId = 0
  AND c.PledgeFlag = 0
  AND c.ContributionTypeId = 1
  AND (@ExcludeFunds = '' OR c.FundId NOT IN
        (SELECT CAST(value AS INT) FROM STRING_SPLIT(@ExcludeFunds, ',') WHERE value <> ''))
GROUP BY f.FamilyId, hoh.Name2, CONVERT(VARCHAR(7), c.ContributionDate, 120)
ORDER BY Household, Month;
