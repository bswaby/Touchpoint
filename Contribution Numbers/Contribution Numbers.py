#Roles=rpt_BusinessAdmin

#sql statement is based on fiscal year.  Ours runs Oct - Sept, so fiscal year is 3 as it's 3 months from the new year. 
#If you want to it to be calendar year, just change it to 0 and update yeartype to nothing or something else you want it to pre-fix

#added a second report called Weekly budget that pulls into a temp table a custom week to week budget to get a over/under comparison. 
#just update values in the temp table to your custom budget

model.Header = 'Contribution Numbers'

fiscalmonth = '3'
yeartype = 'FY'

sqlFundReport = '''
Select
  DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) AS [Year],
  Sum(c.ContributionAmount) AS [Contributed],
  Round(Sum(c.ContributionAmount)/Count(Distinct c.ContributionID),2) AS [AverageGift],
  Count(c.ContributionID) AS [Gifts], 
  Count(Distinct c.PeopleID) AS [UniqueGivers],
  Count(c.ContributionID)/Count(Distinct c.PeopleID) AS [AvgNumofGifts],
  Count(CASE WHEN c.ContributionAmount Between 10000 AND 99999 THEN c.ContributionID END) AS [Gifts10kto99k],
  Count(CASE WHEN c.ContributionAmount >= 100000 THEN c.ContributionID END) AS [Gifts100kPlus]
FROM Contribution c
LEFT JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE Year(c.ContributionDate) >= (DATEPART(Year,getdate())-10)
 Group By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) ---Year(c.ContributionDate) 
 Order By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate)))---Year(c.ContributionDate)
'''
sqlweeklybudget = '''
--Drop Table #tempBudget1
CREATE TABLE #tempBudget1
(
week INT,
budget INT
)

INSERT INTO #tempBudget1
VALUES  (1,230208),
        (2,230208),
        (3,230208),
        (4,230208),
        (5,230208),
        (6,230208),
        (7,230208),
        (8,230208),
        (9,230208),
        (10,300016),
        (11,300000),
        (12,450000),
        (13,900000),
        (14,230208),
        (15,230208),
        (16,230208),
        (17,230208),
        (18,230208),
        (19,230208),
        (20,230208),
        (21,230208),
        (22,230208),
        (23,230208),
        (24,230208),
        (25,230208),
        (26,230208),
        (27,230208),
        (28,230208),
        (29,230208),
        (30,230208),
        (31,230208),
        (32,230208),
        (33,230208),
        (34,230208),
        (35,230208),
        (36,230208),
        (37,230208),
        (38,230208),
        (39,230208),
        (40,230208),
        (41,230208),
        (42,230208),
        (43,230208),
        (44,230208),
        (45,230208),
        (46,230208),
        (47,230208),
        (48,230208),
        (49,230208),
        (50,230208),
        (51,230208),
        (52,230208);

Select
  DATEPART(week,c.ContributionDate) - 40 as [Week],--DATEPART(Year,DATEADD(month,3,(c.ContributionDate))) AS [Year],
  Min(c.ContributionDate) as [WeekStart],

  CASE DATEPART(week,c.ContributionDate) - 40 
    WHEN 1 THEN (Select budget from #tempBudget1 Where week = 1) 
    WHEN 2 THEN (Select budget from #tempBudget1 Where week = 2) 
    WHEN 3 THEN (Select budget from #tempBudget1 Where week = 3) 
    WHEN  4 THEN (Select budget from #tempBudget1 Where week = 4) 
    WHEN  5 THEN (Select budget from #tempBudget1 Where week = 5) 
    WHEN  6 THEN (Select budget from #tempBudget1 Where week = 6) 
    WHEN  7 THEN (Select budget from #tempBudget1 Where week = 7) 
    WHEN  8 THEN (Select budget from #tempBudget1 Where week = 8) 
    WHEN  9 THEN (Select budget from #tempBudget1 Where week = 9) 
    WHEN  10 THEN (Select budget from #tempBudget1 Where week = 10) 
    WHEN  11 THEN (Select budget from #tempBudget1 Where week = 11) 
    WHEN  12 THEN (Select budget from #tempBudget1 Where week = 12) 
    WHEN  13 THEN (Select budget from #tempBudget1 Where week = 13) 
    WHEN  14 THEN (Select budget from #tempBudget1 Where week = 14) 
    WHEN  15 THEN (Select budget from #tempBudget1 Where week = 15) 
    WHEN  16 THEN (Select budget from #tempBudget1 Where week = 16) 
    WHEN  17 THEN (Select budget from #tempBudget1 Where week = 17) 
    WHEN  18 THEN (Select budget from #tempBudget1 Where week = 18) 
    WHEN  19 THEN (Select budget from #tempBudget1 Where week = 19) 
    WHEN  20 THEN (Select budget from #tempBudget1 Where week = 20) 
    WHEN  21 THEN (Select budget from #tempBudget1 Where week = 21) 
    WHEN  22 THEN (Select budget from #tempBudget1 Where week = 22) 
    WHEN  23 THEN (Select budget from #tempBudget1 Where week = 23) 
    WHEN  24 THEN (Select budget from #tempBudget1 Where week = 24) 
    WHEN  25 THEN (Select budget from #tempBudget1 Where week = 25) 
    WHEN  26 THEN (Select budget from #tempBudget1 Where week = 26) 
    WHEN  27 THEN (Select budget from #tempBudget1 Where week = 27) 
    WHEN  28 THEN (Select budget from #tempBudget1 Where week = 28) 
    WHEN  29 THEN (Select budget from #tempBudget1 Where week = 29) 
    WHEN  30 THEN (Select budget from #tempBudget1 Where week = 30) 
    WHEN  31 THEN (Select budget from #tempBudget1 Where week = 31) 
    WHEN  32 THEN (Select budget from #tempBudget1 Where week = 32) 
    WHEN  33 THEN (Select budget from #tempBudget1 Where week = 33) 
    WHEN  34 THEN (Select budget from #tempBudget1 Where week = 34) 
    WHEN  35 THEN (Select budget from #tempBudget1 Where week = 35) 
    WHEN  36 THEN (Select budget from #tempBudget1 Where week = 36) 
    WHEN  37 THEN (Select budget from #tempBudget1 Where week = 37) 
    WHEN  38 THEN (Select budget from #tempBudget1 Where week = 38) 
    WHEN  39 THEN (Select budget from #tempBudget1 Where week = 39) 
    WHEN  40 THEN (Select budget from #tempBudget1 Where week = 40) 
    WHEN  41 THEN (Select budget from #tempBudget1 Where week = 41) 
    WHEN  42 THEN (Select budget from #tempBudget1 Where week = 42) 
    WHEN  43 THEN (Select budget from #tempBudget1 Where week = 43) 
    WHEN  44 THEN (Select budget from #tempBudget1 Where week = 44) 
    WHEN  45 THEN (Select budget from #tempBudget1 Where week = 45) 
    WHEN  46 THEN (Select budget from #tempBudget1 Where week = 46) 
    WHEN  47 THEN (Select budget from #tempBudget1 Where week = 47) 
    WHEN  48 THEN (Select budget from #tempBudget1 Where week = 48) 
    WHEN  49 THEN (Select budget from #tempBudget1 Where week = 49) 
    WHEN  50 THEN (Select budget from #tempBudget1 Where week = 50) 
    WHEN  51 THEN (Select budget from #tempBudget1 Where week = 51) 
    WHEN  52 THEN (Select budget from #tempBudget1 Where week = 52) 
    ELSE 0
    END AS [WklyBudget],

  CASE DATEPART(week,c.ContributionDate) - 40
    WHEN 1 THEN (Select sum(budget) from #tempBudget1 Where week = 1) 
    WHEN 2 THEN (Select sum(budget) from #tempBudget1 Where week <= 2) 
    WHEN 3 THEN (Select sum(budget) from #tempBudget1 Where week <= 3) 
    WHEN  4 THEN (Select sum(budget) from #tempBudget1 Where week <= 4) 
    WHEN  5 THEN (Select sum(budget) from #tempBudget1 Where week <= 5) 
    WHEN  6 THEN (Select sum(budget) from #tempBudget1 Where week <= 6) 
    WHEN  7 THEN (Select sum(budget) from #tempBudget1 Where week <= 7) 
    WHEN  8 THEN (Select sum(budget) from #tempBudget1 Where week <= 8) 
    WHEN  9 THEN (Select sum(budget) from #tempBudget1 Where week <= 9) 
    WHEN  10 THEN (Select sum(budget) from #tempBudget1 Where week <= 10) 
    WHEN  11 THEN (Select sum(budget) from #tempBudget1 Where week <= 11) 
    WHEN  12 THEN (Select sum(budget) from #tempBudget1 Where week <= 12) 
    WHEN  13 THEN (Select sum(budget) from #tempBudget1 Where week <= 13) 
    WHEN  14 THEN (Select sum(budget) from #tempBudget1 Where week <= 14) 
    WHEN  15 THEN (Select sum(budget) from #tempBudget1 Where week <= 15) 
    WHEN  16 THEN (Select sum(budget) from #tempBudget1 Where week <= 16) 
    WHEN  17 THEN (Select sum(budget) from #tempBudget1 Where week <= 17) 
    WHEN  18 THEN (Select sum(budget) from #tempBudget1 Where week <= 18) 
    WHEN  19 THEN (Select sum(budget) from #tempBudget1 Where week <= 19) 
    WHEN  20 THEN (Select sum(budget) from #tempBudget1 Where week <= 20) 
    WHEN  21 THEN (Select sum(budget) from #tempBudget1 Where week <= 21) 
    WHEN  22 THEN (Select sum(budget) from #tempBudget1 Where week <= 22) 
    WHEN  23 THEN (Select sum(budget) from #tempBudget1 Where week <= 23) 
    WHEN  24 THEN (Select sum(budget) from #tempBudget1 Where week <= 24) 
    WHEN  25 THEN (Select sum(budget) from #tempBudget1 Where week <= 25) 
    WHEN  26 THEN (Select sum(budget) from #tempBudget1 Where week <= 26) 
    WHEN  27 THEN (Select sum(budget) from #tempBudget1 Where week <= 27) 
    WHEN  28 THEN (Select sum(budget) from #tempBudget1 Where week <= 28) 
    WHEN  29 THEN (Select sum(budget) from #tempBudget1 Where week <= 29) 
    WHEN  30 THEN (Select sum(budget) from #tempBudget1 Where week <= 30) 
    WHEN  31 THEN (Select sum(budget) from #tempBudget1 Where week <= 31) 
    WHEN  32 THEN (Select sum(budget) from #tempBudget1 Where week <= 32) 
    WHEN  33 THEN (Select sum(budget) from #tempBudget1 Where week <= 33) 
    WHEN  34 THEN (Select sum(budget) from #tempBudget1 Where week <= 34) 
    WHEN  35 THEN (Select sum(budget) from #tempBudget1 Where week <= 35) 
    WHEN  36 THEN (Select sum(budget) from #tempBudget1 Where week <= 36) 
    WHEN  37 THEN (Select sum(budget) from #tempBudget1 Where week <= 37) 
    WHEN  38 THEN (Select sum(budget) from #tempBudget1 Where week <= 38) 
    WHEN  39 THEN (Select sum(budget) from #tempBudget1 Where week <= 39) 
    WHEN  40 THEN (Select sum(budget) from #tempBudget1 Where week <= 40) 
    WHEN  41 THEN (Select sum(budget) from #tempBudget1 Where week <= 41) 
    WHEN  42 THEN (Select sum(budget) from #tempBudget1 Where week <= 42) 
    WHEN  43 THEN (Select sum(budget) from #tempBudget1 Where week <= 43) 
    WHEN  44 THEN (Select sum(budget) from #tempBudget1 Where week <= 44) 
    WHEN  45 THEN (Select sum(budget) from #tempBudget1 Where week <= 45) 
    WHEN  46 THEN (Select sum(budget) from #tempBudget1 Where week <= 46) 
    WHEN  47 THEN (Select sum(budget) from #tempBudget1 Where week <= 47) 
    WHEN  48 THEN (Select sum(budget) from #tempBudget1 Where week <= 48) 
    WHEN  49 THEN (Select sum(budget) from #tempBudget1 Where week <= 49) 
    WHEN  50 THEN (Select sum(budget) from #tempBudget1 Where week <= 50) 
    WHEN  51 THEN (Select sum(budget) from #tempBudget1 Where week <= 51) 
    WHEN  52 THEN (Select sum(budget) from #tempBudget1 Where week <= 52) 
    ELSE 0
    END AS [YTDBudget],
  CASE 
    WHEN DATEPART(week,c.ContributionDate) - 40 = 1 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week = 1) FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) = 1 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 2 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 2)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 2 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 3 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 3)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 3 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 4 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 4)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 4 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 5 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 5)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 5 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 6 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 6)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 6 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 7 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 7)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 7 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 8 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 8)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 8 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 9 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 9)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 9 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 10 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 10)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 10 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 11 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 11)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 11 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 12 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 12)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 12 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 13 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 13)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 13 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 14 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 14)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 14 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 15 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 15)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 15 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 16 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 16)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 16 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 17 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 17)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 17 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 18 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 18)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 18 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 19 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 19)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 19 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 20 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 20)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 20 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 21 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 21)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 21 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 22 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 22)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 22 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 23 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 23)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 23 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 24 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 24)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 24 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 25 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 25)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 25 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 26 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 26)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 26 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 27 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 27)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 27 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 28 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 28)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 28 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 29 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 29)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 29 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 30 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 30)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 30 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 31 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 31)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 31 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 32 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 32)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 32 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 33 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 33)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 33 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 34 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 34)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 34 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 35 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 35)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 35 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 36 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 36)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 36 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 37 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 37)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 37 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 38 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 38)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 38 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 39 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 39)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 39 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 40 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 40)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 40 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 41 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 41)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 41 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 42 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 42)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 42 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 43 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 43)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 43 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 44 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 44)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 44 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 45 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 45)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 45 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 46 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 46)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 46 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 47 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 47)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 47 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 48 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 48)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 48 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 49 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 49)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 49 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 50 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 50)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 50 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 51 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 51)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 51 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 52 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week <= 52)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) <= 52 AND FundID = 1)
    ELSE 0
    END AS [OverUnder],
  
  Sum(c.ContributionAmount) AS [Contributed],
  Round(Sum(c.ContributionAmount)/Count(Distinct c.ContributionID),2) AS [AverageGift],
  Count(c.ContributionID) AS [Gifts], 
  Count(Distinct c.PeopleID) AS [UniqueGivers],
  Count(CASE WHEN c.ContributionAmount Between 10000 AND 99999 THEN c.ContributionID END) AS [Gifts10kto99k],
  Count(CASE WHEN c.ContributionAmount >= 100000 THEN c.ContributionID END) AS [Gifts100kPlus]
FROM Contribution c
LEFT JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE DATEPART(Year,DATEADD(month,3,(c.ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate()))) AND FundID = 1 --(DATEPART(Year,getdate())-10)
 Group By DATEPART(week,c.ContributionDate)--DATEPART(Year,DATEADD(month,3,(c.ContributionDate)))
 Order By DATEPART(week,c.ContributionDate)--DATEPART(Year,DATEADD(month,3,(c.ContributionDate)))

Drop Table #tempBudget1
'''

template = '''
<h3>YoY Contributions</h3>
<table ; width="700px" border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Year</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Contributed</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Avg Gift</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5># of Gifts</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Unique Givers</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Avg # of Gifts</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>10k-99k</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>100k+</h5></b></td>
    </tr>
    {{#each fundreport}}
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><a href="https://myfbch.com/PyScript/ContributionFundYear?p1={{Year}}" target="_blank">''' + yeartype + '''{{Year}}</a></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FmtMoney Contributed}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FmtMoney AverageGift}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt Gifts 'N0'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt UniqueGivers 'N0'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt AvgNumofGifts 'N0'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Gifts10kto99k}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Gifts100kPlus}}</td>
    </tr>
    {{/each}}
    </tbody>
</table>
<br><h3>Weekly Budget</h3>
<table ; width="1000px" border="1" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
    <tbody>
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Week</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Sunday</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Weekly Budget</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>YTD Budget</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Over/Under</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Contributed</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Avg Gift</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>Gifts</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>10k-99k</h5></b></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;"><b><h5>100k+</h5></b></td>
    </tr>
    {{#each weeklybudget}}
    <tr style="{{Bold}}">
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Week}}</a></td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Fmt WeekStart 'd'}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FmtMoney WklyBudget}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FmtMoney YTDBudget}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FmtMoney OverUnder}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FmtMoney Contributed}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{FmtMoney AverageGift}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Gifts}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Gifts10kto99k}}</td>
        <td border: 2px solid #000; vertical-align:top; style="text-align:center;">{{Gifts100kPlus}}</td>
    </tr>
    {{/each}}
    </tbody>
</table>
<br />
<br />
'''


Data.fundreport = q.QuerySql(sqlFundReport)
Data.weeklybudget= q.QuerySql(sqlweeklybudget)
NMReport = model.RenderTemplate(template)
print(NMReport)
