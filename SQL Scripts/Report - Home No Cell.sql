SELECT 
    dbo.People.HomePhone, 
    dbo.People.CellPhone, 
    Organizations_alias1.OrganizationName, 
    dbo.People.Name, 
    dbo.People.FamilyId, 
    dbo.People.PeopleId, 
    dbo.People.PositionInFamilyId 
FROM 
    dbo.ProgDiv 
INNER JOIN 
    dbo.Program 
ON 
    ( 
        dbo.ProgDiv.ProgId = dbo.Program.Id) 
INNER JOIN 
    dbo.Division 
ON 
    ( 
        dbo.ProgDiv.DivId = dbo.Division.Id) 
INNER JOIN 
    dbo.Organizations Organizations_alias1 
ON 
    ( 
        dbo.Division.Id = Organizations_alias1.DivisionId) 
INNER JOIN 
    dbo.OrganizationMembers 
ON 
    ( 
        Organizations_alias1.OrganizationId = dbo.OrganizationMembers.OrganizationId) 
INNER JOIN 
    dbo.People 
ON 
    ( 
        dbo.OrganizationMembers.PeopleId = dbo.People.PeopleId) 
WHERE 
    dbo.Program.Id = @ProgramID 
AND dbo.People.HomePhone <> '' 
AND dbo.People.CellPhone = '' 
ORDER BY 
    dbo.People.HomePhone ASC, 
    dbo.People.CellPhone ASC;