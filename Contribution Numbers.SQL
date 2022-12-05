Select
  Year(c.ContributionDate) AS [Year],
  Sum(c.ContributionAmount) AS [Contributed],
  Round(Sum(c.ContributionAmount)/Count(Distinct c.ContributionID),2) AS [AverageGift],
  Count(c.ContributionID) AS [Gifts], 
  Count(Distinct c.PeopleID) AS [UniqueGivers],
  Count(c.ContributionID)/Count(Distinct c.PeopleID) AS [AvgNumofGifts],
  Count(CASE WHEN c.ContributionAmount Between 10000 AND 99999 THEN c.ContributionID END) AS [Gifts10kto99k],
  Count(CASE WHEN c.ContributionAmount >= 100000 THEN c.ContributionID END) AS [Gifts100kPlus]
FROM Contribution c
INNER JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE Year(c.ContributionDate) >= (DATEPART(Year,getdate())-10)
 Group By Year(c.ContributionDate)
 Order By Year(c.ContributionDate)
