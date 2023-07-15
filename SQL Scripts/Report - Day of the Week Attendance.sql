SELECT  
  DATEPART(week,c.CheckInTime) AS [Week],
  MIN(FORMAT(c.CheckInTime, 'MM-dd-yyyy')) AS [Week Start],
  Count( CASE WHEN DATEPART(dw, c.CheckInTime) = 1 THEN p.PeopleID END) AS [SUNDAY],
  Count( CASE WHEN DATEPART(dw, c.CheckInTime) = 2 THEN p.PeopleID END) AS [MONDAY],
  Count( CASE WHEN DATEPART(dw, c.CheckInTime) = 3 THEN p.PeopleID END) AS [TUESDAY],
  Count( CASE WHEN DATEPART(dw, c.CheckInTime) = 4 THEN p.PeopleID END) AS [WEDNESDAY],
  Count( CASE WHEN DATEPART(dw, c.CheckInTime) = 5 THEN p.PeopleID END) AS [THURSDAY],
  Count( CASE WHEN DATEPART(dw, c.CheckInTime) = 6 THEN p.PeopleID END) AS [FRIDAY],
  Count( CASE WHEN DATEPART(dw, c.CheckInTime) = 7 THEN p.PeopleID END) AS [SATURDAY],
  Count( p.PeopleID) AS [Total]
FROM CheckInTimes c
INNER JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE DATEPART(Year,c.CheckInTime) >= 2023
 Group By DATEPART(week,c.CheckInTime)
 Order By DATEPART(week,c.CheckInTime)