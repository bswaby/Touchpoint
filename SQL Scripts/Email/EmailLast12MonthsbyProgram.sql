SELECT 
    CASE 
        WHEN GROUPING(pro.Name) = 1 THEN 'Total'
        ELSE pro.Name 
    END AS Name,
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -0, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [CurrentMonth],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -1, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [LastMonth],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -2, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [2MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -3, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [3MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -4, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [4MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -5, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [5MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -6, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [6MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -7, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [7MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -8, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [8MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -9, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [9MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -10, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [10MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -11, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [11MonthsAgo],
    SUM(CASE WHEN FORMAT(eqt.Sent, 'yyyy-MM') = FORMAT(DATEADD(month, -12, GETDATE()), 'yyyy-MM') THEN 1 ELSE 0 END) AS [12MonthsAgo]
FROM EmailQueueTo eqt
LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
LEFT JOIN Division d ON d.Id = o.DivisionId
LEFT JOIN Program pro ON pro.Id = d.ProgId
WHERE eqt.Sent >= DATEADD(year, -2, GETDATE())
GROUP BY ROLLUP (pro.Name)
ORDER BY 
    CASE WHEN GROUPING(pro.Name) = 1 THEN 1 ELSE 0 END, 
    pro.Name;
