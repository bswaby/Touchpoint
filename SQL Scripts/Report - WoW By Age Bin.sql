SELECT  
  DATEPART(Week,c.CheckInTime) AS [Week],
    MIN(FORMAT(c.CheckInTime, 'MM-dd-yyyy')) AS [Week Start],
  Count( CASE WHEN p.Age >= 0 AND p.Age < 5 THEN p.PeopleID END) AS [0-5],
  Count( CASE WHEN p.Age >= 6 AND p.Age < 11 THEN p.PeopleID END) AS [6-11],
  Count( CASE WHEN p.Age >= 12 AND p.Age < 17 THEN p.PeopleID END) AS [12-17],
  Count( CASE WHEN p.Age >= 18 AND p.Age < 29 THEN p.PeopleID END) AS [18-29],
  Count( CASE WHEN p.Age >= 30 AND p.Age < 54 THEN p.PeopleID END) AS [30-54],
  Count( CASE WHEN p.Age >= 55 AND p.Age < 64 THEN p.PeopleID END) AS [55-64],
  Count( CASE WHEN p.Age >= 65 THEN p.PeopleID END) AS [65+],
  Count( p.PeopleID) AS [Total]
FROM CheckInTimes c
INNER JOIN
  People p
ON c.PeopleID =  p.PeopleID
WHERE DATEPART(Year,c.CheckInTime) >= 2023
 Group By DATEPART(week,c.CheckInTime)
 Order By DATEPART(week,c.CheckInTime)