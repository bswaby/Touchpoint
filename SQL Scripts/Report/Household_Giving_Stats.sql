/*==============================================================
  HOUSEHOLD GIVING - MEDIAN, MODE, MEAN
  Same filters as the household list. One row of statistics.
==============================================================*/
DECLARE @StartDate DATE = '2025-09-01';
DECLARE @EndDate   DATE = '2026-08-31';
DECLARE @OnlyFunds    VARCHAR(500) = '';
DECLARE @ExcludeFunds VARCHAR(500) = '';
DECLARE @ModeBucket   MONEY = 100;   -- round totals to this before taking the mode

SELECT f.FamilyId, SUM(c.ContributionAmount) AS TotalGiven
INTO #HH
FROM dbo.Contribution c WITH (NOLOCK)
JOIN dbo.People   p ON p.PeopleId = c.PeopleId
JOIN dbo.Families f ON f.FamilyId = p.FamilyId
WHERE c.ContributionDate >= @StartDate
  AND c.ContributionDate <  DATEADD(DAY, 1, @EndDate)
  AND c.ContributionStatusId = 0
  AND c.PledgeFlag = 0
  AND c.ContributionTypeId = 1
  AND (@OnlyFunds = '' OR c.FundId IN
        (SELECT CAST(value AS INT) FROM STRING_SPLIT(@OnlyFunds, ',') WHERE value <> ''))
  AND (@ExcludeFunds = '' OR c.FundId NOT IN
        (SELECT CAST(value AS INT) FROM STRING_SPLIT(@ExcludeFunds, ',') WHERE value <> ''))
GROUP BY f.FamilyId;

SELECT
    COUNT(*)                                              AS GivingHouseholds,
    SUM(TotalGiven)                                       AS TotalGiven,
    CAST(AVG(TotalGiven) AS DECIMAL(12,2))                AS MeanHousehold,
    (SELECT DISTINCT PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY TotalGiven) OVER () FROM #HH)
                                                          AS MedianHousehold,
    (SELECT DISTINCT PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY TotalGiven) OVER () FROM #HH)
                                                          AS P25,
    (SELECT DISTINCT PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY TotalGiven) OVER () FROM #HH)
                                                          AS P75,
    (SELECT TOP 1 ROUND(TotalGiven / @ModeBucket, 0) * @ModeBucket
       FROM #HH GROUP BY ROUND(TotalGiven / @ModeBucket, 0) * @ModeBucket
      ORDER BY COUNT(*) DESC, 1)                          AS ModeBucket,
    MIN(TotalGiven)                                       AS Smallest,
    MAX(TotalGiven)                                       AS Largest
FROM #HH;

DROP TABLE #HH;
