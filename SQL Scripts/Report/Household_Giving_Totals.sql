/*==============================================================
  HOUSEHOLD GIVING TOTALS
  One row per family, total given in a date range.

  Edit the settings block, run, then Export to Excel.
==============================================================*/

---------------- SETTINGS ----------------
DECLARE @StartDate DATE = '2025-09-01';
DECLARE @EndDate   DATE = '2026-08-31';   -- inclusive

-- Leave blank for all funds. Otherwise a comma list of FundIds.
DECLARE @OnlyFunds    VARCHAR(500) = '';   -- e.g. '1,3,7'  (whitelist)
DECLARE @ExcludeFunds VARCHAR(500) = '';   -- e.g. '42,55'  (blacklist)

-- Count non-tax-deductible gifts, gift-in-kind and stock?
DECLARE @IncludeNonTaxDeductible BIT = 0;  -- type 9
DECLARE @IncludeGiftInKind       BIT = 0;  -- type 10
DECLARE @IncludeStock            BIT = 0;  -- type 20

-- Hide families under this total (0 shows everyone who gave)
DECLARE @MinTotal MONEY = 0;
------------------------------------------

DECLARE @Types TABLE (Id INT PRIMARY KEY);
INSERT INTO @Types (Id) VALUES (1);                                   -- Tax deductible
IF @IncludeNonTaxDeductible = 1 INSERT INTO @Types (Id) VALUES (9);
IF @IncludeGiftInKind       = 1 INSERT INTO @Types (Id) VALUES (10);
IF @IncludeStock            = 1 INSERT INTO @Types (Id) VALUES (20);
-- Never counted: 6 Returned Check, 7 Reversed, 8 Pledge, 99 Non-contribution

SELECT
    f.FamilyId,
    ISNULL(hoh.Name2, MIN(p.Name2))                    AS Household,
    ISNULL(hoh.LastName, MIN(p.LastName))              AS FamilyName,
    COUNT(*)                                           AS Gifts,
    COUNT(DISTINCT c.PeopleId)                         AS Givers,
    SUM(c.ContributionAmount)                          AS TotalGiven,
    CAST(AVG(c.ContributionAmount) AS DECIMAL(12,2))   AS AvgGift,
    MIN(c.ContributionDate)                            AS FirstGift,
    MAX(c.ContributionDate)                            AS LastGift
FROM dbo.Contribution c WITH (NOLOCK)
JOIN dbo.People   p ON p.PeopleId  = c.PeopleId
JOIN dbo.Families f ON f.FamilyId  = p.FamilyId
LEFT JOIN dbo.People hoh ON hoh.PeopleId = f.HeadOfHouseholdId
WHERE c.ContributionDate >= @StartDate
  AND c.ContributionDate <  DATEADD(DAY, 1, @EndDate)
  AND c.ContributionStatusId = 0                 -- Recorded only
  AND c.PledgeFlag = 0                           -- a pledge is a promise
  AND c.ContributionTypeId IN (SELECT Id FROM @Types)
  AND (@OnlyFunds = '' OR c.FundId IN
        (SELECT CAST(value AS INT) FROM STRING_SPLIT(@OnlyFunds, ',') WHERE value <> ''))
  AND (@ExcludeFunds = '' OR c.FundId NOT IN
        (SELECT CAST(value AS INT) FROM STRING_SPLIT(@ExcludeFunds, ',') WHERE value <> ''))
GROUP BY f.FamilyId, hoh.Name2, hoh.LastName
HAVING SUM(c.ContributionAmount) >= @MinTotal
ORDER BY TotalGiven DESC;
