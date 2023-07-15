SELECT  
  DATEPART(week,c.CheckInTime) AS [Week],
  MIN(FORMAT(c.CheckInTime, 'MM-dd-yyyy')) AS [Week Start],
  Count( CASE WHEN DATEPART(hour,c.CheckInTime) >= 4 AND DATEPART(hour,c.CheckInTime) <= 7 THEN p.PeopleID END) AS [4AM-7AM],
  Count( CASE WHEN DATEPART(hour,c.CheckInTime) >= 8 AND DATEPART(hour,c.CheckInTime) <= 11 THEN p.PeopleID END) AS [8AM-11AM],
  Count( CASE WHEN DATEPART(hour,c.CheckInTime) >= 12 AND DATEPART(hour,c.CheckInTime) <= 15 THEN p.PeopleID END) AS [12AM-3PM],
  Count( CASE WHEN DATEPART(hour,c.CheckInTime) >= 16 AND DATEPART(hour,c.CheckInTime) <= 19 THEN p.PeopleID END) AS [4PM-7PM],
  Count( p.PeopleID) AS [Total]
FROM CheckInTimes c
INNER JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE DATEPART(Year,c.CheckInTime) >= 2023
 Group By DATEPART(week,c.CheckInTime)
 Order By DATEPART(week,c.CheckInTime)