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
IF OBJECT_ID('tempdb.dbo.#tempCalendar1', 'U') IS NOT NULL
	DROP TABLE #tempCalendar1
IF OBJECT_ID('tempdb.dbo.#tempBudget1', 'U') IS NOT NULL
	DROP TABLE #tempBudget1

DECLARE @StartDate  date = '20230101';

DECLARE @CutoffDate date = DATEADD(DAY, -1, DATEADD(YEAR, 1, @StartDate));

;WITH seq(n) AS 
(
  SELECT 0 UNION ALL SELECT n + 1 FROM seq
  WHERE n < DATEDIFF(DAY, @StartDate, @CutoffDate)
),
d(d) AS 
(
  SELECT DATEADD(DAY, n, @StartDate) FROM seq
),
src AS
(
  SELECT
    TheDate         = CONVERT(date,		  Dateadd(month, -3, d)),
    TheDay          = DATEPART(DAY,       d),
    TheDayName      = DATENAME(WEEKDAY,   d),
    TheWeek         = DATEPART(WEEK,      d),
    TheISOWeek      = DATEPART(ISO_WEEK,  d),
    TheDayOfWeek    = DATEPART(WEEKDAY,   d),
    TheMonth        = DATEPART(MONTH,     d),
    TheMonthName    = DATENAME(MONTH,     d),
    TheQuarter      = DATEPART(Quarter,   d),
    TheYear         = DATEPART(YEAR,      d),
    TheFirstOfMonth = DATEFROMPARTS(YEAR(Dateadd(month, -3, d)), MONTH(Dateadd(month, -3, d)), 1),
    TheLastOfYear   = DATEFROMPARTS(YEAR(d), 12, 31),
    TheDayOfYear    = DATEPART(DAYOFYEAR, d)
  FROM d
),
dim AS
(
  SELECT
	--DATEADD(day, 7 * (Number - 1), @startDate)Number,
    TheDate, 
    TheDay,
    TheDaySuffix        = CONVERT(char(2), CASE WHEN TheDay / 10 = 1 THEN 'th' ELSE 
                            CASE RIGHT(TheDay, 1) WHEN '1' THEN 'st' WHEN '2' THEN 'nd' 
                            WHEN '3' THEN 'rd' ELSE 'th' END END),
    TheDayName,
    TheDayOfWeek,
    TheDayOfWeekInMonth = CONVERT(tinyint, ROW_NUMBER() OVER 
                            (PARTITION BY TheFirstOfMonth, TheDayOfWeek ORDER BY TheDate)),
    TheDayOfYear,
    IsWeekend           = CASE WHEN TheDayOfWeek IN (CASE @@DATEFIRST WHEN 1 THEN 6 WHEN 7 THEN 1 END,7) 
                            THEN 1 ELSE 0 END,
    TheWeek,
    TheISOweek,
    --TheFirstOfWeek      = DATEADD(DAY, TheDayOfWeek - 1, TheDate),
	TheFirstOfWeek = CAST(DATEADD(DAY, -1*(DATEPART(WEEKDAY,TheDate)-1),TheDate) as Date),
    TheLastOfWeek       = DATEADD(DAY, 6, DATEADD(DAY, 1 - TheDayOfWeek, TheDate)),
    TheWeekOfMonth      = CONVERT(tinyint, DENSE_RANK() OVER 
                            (PARTITION BY TheYear, TheMonth ORDER BY TheWeek)),
    TheMonth,
    TheMonthName,
    TheFirstOfMonth,
    TheLastOfMonth      = MAX(TheDate) OVER (PARTITION BY TheYear, TheMonth),
    TheFirstOfNextMonth = DATEADD(MONTH, 1, TheFirstOfMonth),
    TheLastOfNextMonth  = DATEADD(DAY, -1, DATEADD(MONTH, 2, TheFirstOfMonth)),
    TheQuarter,
    TheFirstOfQuarter   = MIN(TheDate) OVER (PARTITION BY TheYear, TheQuarter),
    TheLastOfQuarter    = MAX(TheDate) OVER (PARTITION BY TheYear, TheQuarter),
    TheYear,
    TheISOYear          = TheYear - CASE WHEN TheMonth = 1 AND TheISOWeek > 51 THEN 1 
                            WHEN TheMonth = 12 AND TheISOWeek = 1  THEN -1 ELSE 0 END,      
    TheFirstOfYear      = DATEFROMPARTS(TheYear, 1,  1),
    TheLastOfYear,
    IsLeapYear          = CONVERT(bit, CASE WHEN (TheYear % 400 = 0) 
                            OR (TheYear % 4 = 0 AND TheYear % 100 <> 0) 
                            THEN 1 ELSE 0 END),
    Has53Weeks          = CASE WHEN DATEPART(WEEK,     TheLastOfYear) = 53 THEN 1 ELSE 0 END,
    Has53ISOWeeks       = CASE WHEN DATEPART(ISO_WEEK, TheLastOfYear) = 53 THEN 1 ELSE 0 END,
    MMYYYY              = CONVERT(char(2), CONVERT(char(8), TheDate, 101))
                          + CONVERT(char(4), TheYear),
    Style101            = CONVERT(char(10), TheDate, 101),
    Style103            = CONVERT(char(10), TheDate, 103),
    Style112            = CONVERT(char(8),  TheDate, 112),
    Style120            = CONVERT(char(10), TheDate, 120)
  FROM src
)

SELECT * INTO #tempCalendar1 FROM dim
ORDER BY TheDate
OPTION (MAXRECURSION 0);

CREATE TABLE #tempBudget1
(
week DATE,
budget INT
)

INSERT INTO #tempBudget1
VALUES  ('09/26/2021',230208),
('10/03/2021',213962),
('10/10/2021',213962),
('10/17/2021',213962),
('10/24/2021',213962),
('10/31/2021',213962),
('11/07/2021',213962),
('11/14/2021',213962),
('11/21/2021',213962),
('11/28/2021',213962),
('12/05/2021',503833),
('12/12/2021',213962),
('12/19/2021',412137),
('12/26/2021',532453),
('12/31/2021',699982),
('01/02/2022',199860),
('01/09/2022',199860),
('01/16/2022',199860),
('01/23/2022',199860),
('01/30/2022',199860),
('02/06/2022',199860),
('02/13/2022',199860),
('02/20/2022',199860),
('02/27/2022',199860),
('03/06/2022',199860),
('03/13/2022',199860),
('03/20/2022',199860),
('03/27/2022',199860),
('04/03/2022',199860),
('04/10/2022',199860),
('04/17/2022',199860),
('04/24/2022',199860),
('05/01/2022',199860),
('05/08/2022',199860),
('05/15/2022',199860),
('05/22/2022',199860),
('05/29/2022',199860),
('06/05/2022',199860),
('06/12/2022',199860),
('06/19/2022',199860),
('06/26/2022',199860),
('07/03/2022',199860),
('07/10/2022',199860),
('07/17/2022',199860),
('07/24/2022',199860),
('07/31/2022',199860),
('08/07/2022',199860),
('08/14/2022',199860),
('08/21/2022',199860),
('08/28/2022',199860),
('09/04/2022',199860),
('09/11/2022',199860),
('09/18/2022',199860),
('09/25/2022',199860),
('10/02/2022',230208),
('10/09/2022',230208),
('10/16/2022',230208),
('10/23/2022',230208),
('10/30/2022',230208),
('11/06/2022',230208),
('11/13/2022',230208),
('11/20/2022',230208),
('11/27/2022',230208),
('12/04/2022',300016),
('12/11/2022',300000),
('12/18/2022',450000),
('12/25/2022',450000),
('12/31/2022',450000),
('01/01/2023',230208),
('01/08/2023',230208),
('01/15/2023',230208),
('01/22/2023',230208),
('01/29/2023',230208),
('02/05/2023',230208),
('02/12/2023',230208),
('02/19/2023',230208),
('02/26/2023',230208),
('03/05/2023',230208),
('03/12/2023',230208),
('03/19/2023',230208),
('03/26/2023',230208),
('04/02/2023',230208),
('04/09/2023',230208),
('04/16/2023',230208),
('04/23/2023',230208),
('04/30/2023',230208),
('05/07/2023',230208),
('05/14/2023',230208),
('05/21/2023',230208),
('05/28/2023',230208),
('06/04/2023',230208),
('06/11/2023',230208),
('06/18/2023',230208),
('06/25/2023',230208),
('07/02/2023',230208),
('07/09/2023',230208),
('07/16/2023',230208),
('07/23/2023',230208),
('07/30/2023',230208),
('08/06/2023',230208),
('08/13/2023',230208),
('08/20/2023',230208),
('08/27/2023',230208),
('09/03/2023',230208),
('09/10/2023',230208),
('09/17/2023',230208),
('09/24/2023',230208)

Select --Distinct 
  tc.TheFirstofWeek 
  ,tb.budget as [WklyBudget]
  ,CASE WHEN tc.TheFirstofWeek <= tb.week 
    THEN (Select SUM(budget) FROM #tempBudget1 WHERE week Between '20221002' AND tc.TheFirstOfWeek) END AS [YTDBudget]
  ,CASE WHEN tb.budget > 0
    THEN (Select SUM(c1.ContributionAmount) FROM Contribution c1 WHERE (c1.ContributionDate Between '20221002' AND tc.TheFirstOfWeek) and c1.FundId = 1)
	END AS [TotalGiving]
  ,SUM(c.ContributionAmount) AS Contributed
  ,CASE WHEN tb.budget > 0
    THEN ((Select SUM(c1.ContributionAmount) FROM Contribution c1 WHERE (c1.ContributionDate Between '20221002' AND tc.TheFirstOfWeek) and c1.FundId = 1)
	- (Select SUM(tb1.budget) FROM #tempBudget1 tb1 WHERE tb1.week Between '20221002' AND tc.TheFirstOfWeek))
	END AS [OverUnder]
  ,DATEADD(WEEK,-52,tc.TheFirstofWeek) AS [PreviousYear]
  ,CASE WHEN tb.budget > 0
    THEN (SELECT SUM(c1.ContributionAmount) FROM Contribution c1 WHERE c1.ContributionDate = DATEADD(WEEK,-52,tc.TheFirstOfWeek) and c1.FundId = 1) 
	END AS pyContributed
  ,CASE WHEN tc.TheFirstofWeek <= tb.week 
    THEN (Select SUM(budget) FROM #tempBudget1 WHERE week Between '20211003' AND DATEADD(WEEK,-52,tc.TheFirstofWeek)) END AS [pyYTDBudget]
  ,CASE WHEN tc.TheFirstofWeek <= tb.week 
    THEN (Select SUM(budget) FROM #tempBudget1 WHERE week Between '20211003' AND DATEADD(WEEK,-52,tc.TheFirstofWeek)) END AS [pyYTDBudget]
  ,CASE WHEN tb.budget > 0
    THEN (Select SUM(c1.ContributionAmount) FROM Contribution c1 WHERE (c1.ContributionDate Between '20211003' AND DATEADD(WEEK,-52,tc.TheFirstOfWeek)) and c1.FundId = 1)
	END AS [pyTotalGiving]
  ,CASE WHEN tb.budget > 0 
    THEN ((Select SUM(c2.ContributionAmount) FROM Contribution c2 WHERE (c2.ContributionDate Between '20211003' AND DATEADD(WEEK,-52,tc.TheFirstOfWeek)) and c2.FundId = 1)
	- (Select SUM(tb1.budget) FROM #tempBudget1 tb1 WHERE tb1.week Between '20211003' AND DATEADD(WEEK,-52,tc.TheFirstofWeek)))
	END AS [pyOverUnder]  
  ,Round(Sum(c.ContributionAmount)/Count(Distinct c.ContributionID),2) AS [AverageGift]
  ,Count(c.ContributionID) AS [Gifts]
  ,Count(Distinct c.PeopleID) AS [UniqueGivers]
  ,Count(CASE WHEN c.ContributionAmount Between 10000 AND 99999 THEN c.ContributionID END) AS [Gifts10kto99k]
  ,Count(CASE WHEN c.ContributionAmount >= 100000 THEN c.ContributionID END) AS [Gifts100kPlus]
FROM #tempBudget1 tb
LEFT JOIN #tempCalendar1 tc ON tb.week = tc.TheFirstOfWeek
LEFT JOIN Contribution c ON c.ContributionDate = tc.TheDate
WHERE c.FundId = 1
--WHERE tc.TheFirstOfWeek <> '20220925'
GROUP BY
  tc.TheFirstOfWeek
  ,tb.budget
  ,tb.week
Order By tc.TheFirstOfWeek

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
              data.addColumn('number', 'pyContributed');
        
              data.addRows([
                {{#each contributiondata}}
                  [
                    '{{Fmt TheFirstofWeek 'd'}}', 
                    {{WklyBudget}}, 
                    {{Contributed}},
                    {{pyContributed}},
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
                width: 900,
                height: 500,
                axes: {
                  x: {
                    0: {side: 'top'}
                  }
                }   
              };
     
              var chart = new google.visualization.LineChart(document.getElementById('chart_div'));
              chart.draw(data, options);

          }
    </script>
    <div id='chart_div' style='width: 900px; height: 500px;'></div>

    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
      google.charts.load('current', {'packages':['table']});
      google.charts.setOnLoadCallback(drawTable);

      function drawTable() {
        var data = new google.visualization.DataTable();
        data.addColumn('string', 'Sunday');
        data.addColumn('number', 'Weekly Budget');
        data.addColumn('number', 'YTD Budget');
        data.addColumn('number', 'Contributed');
        data.addColumn('number', 'OverUnder');
        data.addColumn('number', 'Total Giving');
        data.addColumn('number', 'pyContributed');
        data.addColumn('number', 'pyOverUnder');
        data.addColumn('number', 'pyTotal Giving');
        data.addColumn('number', 'Avg Gift');
        data.addColumn('number', 'Gifts');
        data.addColumn('number', '10k-99k');
        data.addColumn('number', '100k+');
        
        data.addRows([
            {{#each contributiondata}}
                [
                '<a href="https://myfbch.com/PyScript/ContributionWeekReport?P1=%27{{Fmt TheFirstofWeek 'd'}}%27" target="_blank">{{Fmt TheFirstofWeek 'd'}}</a>',
                {v: {{WklyBudget}},  f: '{{FmtMoney WklyBudget}}'},
                {v: {{YTDBudget}},  f: '{{FmtMoney YTDBudget}}'},
                {v: {{Contributed}},  f: '{{FmtMoney Contributed}}'},
                {v: {{OverUnder}},  f: '{{FmtMoney OverUnder}}'},
                {v: {{TotalGiving}},  f: '{{FmtMoney TotalGiving}}'},
                {v: {{pyContributed}},  f: '{{FmtMoney pyContributed}}'},
                {v: {{pyOverUnder}},  f: '{{FmtMoney pyOverUnder}}'},
                {v: {{pyTotalGiving}},  f: '{{FmtMoney pyTotalGiving}}'},
                {v: {{AverageGift}},  f: '{{FmtMoney AverageGift}}'},
                {{Gifts}},
                {{Gifts10kto99k}},
                {{Gifts100kPlus}}
                ],
            {{/each}}

        ]);
        
        var table = new google.visualization.Table(document.getElementById('table_wklybudget'));
        table.draw(data, {showRowNumber: false, alternatingRowStyle: true, allowHtml: true, width: '100%', height: '100%'});
      }
    </script>
    <div id='table_wklybudget' style='width: 1100px; height: 600px;'></div>
'''

Data.fundreport = q.QuerySql(sqlFundReport)
Data.contributiondata = q.QuerySql(sqlcontributiondata)
NMReport = model.RenderTemplate(template)
print(NMReport)
