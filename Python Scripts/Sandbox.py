programNameQuery =  '''
SELECT Name
FROM Program
WHERE Name = 'FMC'
'''
for a in q.QuerySql(programNameQuery):
    programName = a.Name
    
print programName