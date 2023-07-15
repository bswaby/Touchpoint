#Roles=rpt_BusinessAdmin

#sql statement is based on fiscal year.  Ours runs Oct - Sept, so fiscal year is 3 as it's 3 months from the new year. 
#If you want to it to be calendar year, just change it to 0 and update yeartype to nothing or something else you want it to pre-fix
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
 Order By DATEPART(Year,DATEADD(month,''' + fiscalmonth + ''',(c.ContributionDate))) Desc---Year(c.ContributionDate)
'''
sqlcontributiondata = '''
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
  --DATEPART(week,c.ContributionDate) AS [ACTUALWEEK],
  CASE WHEN DATEPART(week,c.ContributionDate) BETWEEN 1 AND 40 
   THEN DATEPART(week,c.ContributionDate) + 13
   ELSE DATEPART(week,c.ContributionDate) - 40 END AS [Week],
  -- DATEPART(week,c.ContributionDate) - 40 as [Week],--DATEPART(Year,DATEADD(month,3,(c.ContributionDate))) AS [Year],
  Min(c.ContributionDate) as [WeekStart],

  CASE WHEN DATEPART(week,c.ContributionDate) BETWEEN 1 AND 40
    THEN (Select budget from #tempBudget1 Where week = (DATEPART(week,c.ContributionDate) + 13)) 
    ELSE (Select budget from #tempBudget1 Where week = (DATEPART(week,c.ContributionDate) - 40))
    END AS [WklyBudget],

  CASE WHEN DATEPART(week,c.ContributionDate) BETWEEN 1 AND 40
    THEN (Select sum(budget) from #tempBudget1 Where week <= (DATEPART(week,c.ContributionDate) + 13)) 
    ELSE (Select sum(budget) from #tempBudget1 Where week <= (DATEPART(week,c.ContributionDate) - 40))
    END AS [YTDBudget],
    
  CASE 
    WHEN DATEPART(week,c.ContributionDate) - 40 = 1 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week = 1) FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) = 1 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 2 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND  2)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND 2 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 3 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND  3)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND 3 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 4 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 4)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  4 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 5 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 5)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  5 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 6 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 6)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  6 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 7 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 7)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  7 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 8 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 8)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  8 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 9 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 9)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  9 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 10 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 10)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  10 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 11 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 11)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  11 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 12 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 12)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  12 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 13 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 13)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  13 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 14 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 14)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  14 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 15 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 15)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  15 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 16 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 16)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  16 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 17 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 17)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  17 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 18 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 18)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  18 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 19 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 19)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  19 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) - 40 <= 20 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 20)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  20 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 21 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 21)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  21 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 22 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 22)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  22 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 23 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 23)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  23 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 24 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 24)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  24 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 25 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 25)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  25 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 26 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 26)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  26 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 27 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 27)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  27 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 28 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 28)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  28 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 29 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 29)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  29 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 30 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 30)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  30 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 31 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 31)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  31 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 32 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 32)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  32 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 33 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 33)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  33 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 34 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 34)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  34 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 35 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 35)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  35 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 36 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 36)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  36 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 37 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 37)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  37 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 38 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 38)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  38 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 39 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 39)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) + 13) BETWEEN 1 AND  39 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 40 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 40)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  40 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 41 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 41)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  41 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 42 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 42)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  42 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 43 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 43)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  43 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 44 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 44)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  44 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 45 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 45)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  45 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 46 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 46)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  46 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 47 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 47)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  47 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 48 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 48)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  48 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 49 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 49)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  49 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 50 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 50)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  50 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 51 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 51)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  51 AND FundID = 1)
    WHEN DATEPART(week,c.ContributionDate) <= 52 THEN (Select Sum(ContributionAmount) - (Select sum(budget) from #tempBudget1 Where week BETWEEN 1 AND 52)  FROM Contribution
                                                                    WHERE DATEPART(Year,DATEADD(month,3,(ContributionDate))) >= DATEPART(Year,DATEADD(month,3,(getdate())))
                                                                      AND (DATEPART(week,ContributionDate) - 40) BETWEEN 1 AND  52 AND FundID = 1)
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
 Order By Week Desc, DATEPART(week,c.ContributionDate) Desc --DATEPART(Year,DATEADD(month,3,(c.ContributionDate))) 

Drop Table #tempBudget1
'''

template = '''
    <h3>YoY Contributions</h3>
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">

      google.charts.load('current', {'packages':['table']});
      google.charts.setOnLoadCallback(drawTable);

      function drawTable() {
        var data = new google.visualization.DataTable();
        data.addColumn('string', 'Year');
        data.addColumn('number', 'Contributed');
        data.addColumn('number', 'Avg Gift');
        data.addColumn('number', 'Gifts');
        data.addColumn('number', 'Unique Givers');
        data.addColumn('number', 'Avg Gifts');
        data.addColumn('number', '10-99k');
        data.addColumn('number', '100k+');

        data.addRows([
            {{#each fundreport}}
                [
                '<a href="https://myfbch.com/PyScript/ContributionFundYear?p1={{Year}}" target="_blank">{{Year}}</a>',
                {v: {{Contributed}},  f: '{{FmtMoney Contributed}}'}, 
                {v: {{AverageGift}},  f: '{{FmtMoney AverageGift}}'}, 
                {v: {{Gifts}},  f: '{{Fmt Gifts 'N0'}}'},
                {v: {{UniqueGivers}},  f: '{{Fmt UniqueGivers 'N0'}}'},
                {v: {{AvgNumofGifts}},  f: '{{Fmt AverageGift 'N0'}}'},
                {{Gifts10kto99k}},
                {{Gifts100kPlus}}
                ],
            {{/each}}
        ]);
        
        var table = new google.visualization.Table(document.getElementById('table_div'));
        table.draw(data, {showRowNumber: false, alternatingRowStyle: true, allowHtml: true, width: '100%', height: '100%'});
      }
    </script>
    <div id='table_div' style='width: 550px; height: 320px;'></div>
    
    <br><h3>Weekly Budget</h3>
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
        google.charts.load('current', {packages: ['corechart', 'line']});
        google.charts.setOnLoadCallback(drawTrendlines);
        
        function drawTrendlines() {
              var data = new google.visualization.DataTable();
              data.addColumn('string', 'Sunday')
              data.addColumn('number', 'budget');
              data.addColumn('number', 'contributed');
        
              data.addRows([
                {{#each contributiondata}}
                  [
                    '{{Fmt WeekStart 'd'}}', 
                    {{WklyBudget}}, 
                    {{Contributed}},
                  ],
                {{/each}}    
              ]);
        
              var options = {
                hAxis: {
                  title: 'Fiscal Week'
                },
                vAxis: {
                  title: 'Amount'
                },
                colors: ['#AB0D06', '#007329'],

              };
        
              var chart = new google.visualization.LineChart(document.getElementById('chart_div'));
              chart.draw(data, options);
          }
    </script>
    <div id='chart_div' style='width: 800px; height: 300px;'></div>

    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
      google.charts.load('current', {'packages':['table']});
      google.charts.setOnLoadCallback(drawTable);

      function drawTable() {
        var data = new google.visualization.DataTable();
        data.addColumn('number', 'Week');
        data.addColumn('string', 'Sunday');
        data.addColumn('number', 'Weekly Budget');
        data.addColumn('number', 'YTD Budget');
        data.addColumn('number', 'OverUnder');
        data.addColumn('number', 'Contributed');
        data.addColumn('number', 'Avg Gift');
        data.addColumn('number', 'Gifts');
        data.addColumn('number', '10k-99k');
        data.addColumn('number', '100k+');
        
        data.addRows([
            {{#each contributiondata}}
                [
                {{Week}}, 
                '{{Fmt WeekStart 'd'}}', 
                {v: {{WklyBudget}},  f: '{{FmtMoney WklyBudget}}'}, 
                {v: {{YTDBudget}},  f: '{{FmtMoney YTDBudget}}'}, 
                {v: {{OverUnder}},  f: '{{FmtMoney OverUnder}}'},
                {v: {{Contributed}},  f: '{{FmtMoney Contributed}}'},
                {v: {{AverageGift}},  f: '{{FmtMoney AverageGift}}'},
                {{Gifts}},
                {{Gifts10kto99k}},
                {{Gifts100kPlus}}
                ],
            {{/each}}

        ]);
        
        var table = new google.visualization.Table(document.getElementById('table_wklybudget'));
        table.draw(data, {showRowNumber: false, alternatingRowStyle: true, width: '100%', height: '100%'});
      }
    </script>
    <div id='table_wklybudget' style='width: 800px; height: 400px;'></div>
'''

Data.fundreport = q.QuerySql(sqlFundReport)
Data.contributiondata = q.QuerySql(sqlcontributiondata)
NMReport = model.RenderTemplate(template)
print(NMReport)
