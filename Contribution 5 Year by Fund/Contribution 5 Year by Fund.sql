SELECT 
    Distinct dbo.ContributionFund.FundName, 
--    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2013 THEN dbo.Contribution.ContributionAmount END) AS [2013], 
--    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2014 THEN dbo.Contribution.ContributionAmount END) AS [2014], 
--    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2015 THEN dbo.Contribution.ContributionAmount END) AS [2015], 
--    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2016 THEN dbo.Contribution.ContributionAmount END) AS [2016], 
--    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2017 THEN dbo.Contribution.ContributionAmount END) AS [2017], 
    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2018 THEN dbo.Contribution.ContributionAmount END) AS [2018], 
    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2019 THEN dbo.Contribution.ContributionAmount END) AS [2019],     
    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2020 THEN dbo.Contribution.ContributionAmount END) AS [2020], 
    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2021 THEN dbo.Contribution.ContributionAmount END) AS [2021],
    SUM(CASE WHEN Year(dbo.Contribution.ContributionDate) = 2022 THEN dbo.Contribution.ContributionAmount END) AS [2022] 
FROM 
    dbo.Contribution 
INNER JOIN 
    dbo.ContributionFund 
ON 
    ( 
        dbo.Contribution.FundId = dbo.ContributionFund.FundId) 
Where dbo.ContributionFund.FundName is not null AND Year(dbo.Contribution.ContributionDate) Between 2018 AND 2022
GROUP BY 
    dbo.ContributionFund.FundName 
