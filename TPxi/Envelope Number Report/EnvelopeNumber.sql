-- ---------------------------------------------------------------
-- Envelope Number Report
-- ---------------------------------------------------------------
-- Written By: Ben Swaby (TPxi Software, LLC)
-- Email: bswaby@fbchtn.org                                                                                                      
-- Website: https://tpxisoftware.com
-- GitHub: https://github.com/bswaby/Touchpoint  (50+ free tools)                                                                
-- ----------------------------------------------------------------                                                              
-- These tools are free because they should be.
-- If they've saved you time or helped your team, and you want to                                                                
-- support continued development, check out:                                                                                     
--
-- DisplayCache(TM) - church digital signage that integrates with TouchPoint(R)                                                  
-- https://displaycache.com                                
--
-- TPxi Go(TM) - your church contacts, wherever you work.
-- Look up anyone in TouchPoint(R), log calls and emails from Outlook                                                            
-- or your phone. No tab switching, no lost context.
-- https://tpxigo.com                                                                                                            
-- ----------------------------------------------------------------
-- Description:
--   Generates a mailing list of all envelope holders with their
--   envelope number, name (formatted for joint/individual/deceased
--   spouse), and full mailing address. Handles joint envelopes
--   (Mr. and Mrs.), individual envelopes, deceased spouse fallback,
--   and flags children under 18 with envelope options set.
--
-- Assumptions:
--   - Envelope numbers are stored as an Extra Value (PeopleExtra)
--     with Field = 'EnvelopeNumber' and IntValue
--   - EnvelopeOptionsId: 1 = Individual, 2 = Joint
--   - Excludes deceased, archived, no-address, and do-not-mail
--
-- Setup:
--   1. Admin > Advanced > Special Content > SQL Scripts
--   2. New SQL Script File, name it "Envelope Number Report"
--   3. Paste this SQL
--   4. Add to CustomReports:
--      <Report name="Envelope Number Report" type="SqlReport" role="Finance" />
-- ---------------------------------------------------------------

;With EnvNum as (
Select pe.PeopleId
, pe.Field
, pe.IntValue
From PeopleExtra pe
LEFT Join People pp ON pp.PeopleId = pe.PeopleId
Inner Join Families fam ON pp.FamilyId = fam.FamilyId
Where (fam.HeadOfHouseholdId = (pp.PeopleId))
AND Field = 'EnvelopeNumber'
--AND pp.PositionInFamilyId = 10
AND pp.EnvelopeOptionsId IN (1,2)
AND (NOT (pp.IsDeceased = 1)) 
AND (NOT (pp.ArchivedFlag = 1))
AND pp.PrimaryAddress <> ''
AND (NOT (pp.DoNotMailFlag = 1))),

EnvJoint AS(
Select 
  CASE WHEN p.EnvelopeOptionsId = 2 THEN CONCAT(p.TitleCode, ' and ', sp.TitleCode, ' ', p.FirstName, ' ', p.LastName) ELSE CONCAT(p.TitleCode, ' ', p.FirstName, ' ', p.LastName) END AS [Name]--, ' and ', s.TitleCode, ' ', s.FirstName, ' ', p.LastName) AS [Name]
, p.PrimaryAddress
, p.PrimaryAddress2
, p.PrimaryCity
, p.PrimaryState
, p.PrimaryZip
, CASE WHEN en.IntValue is NULL THEN cast(p.PeopleId as varchar) ELSE CONCAT(cast(en.IntValue AS varchar), '-', cast(p.PeopleId AS varchar)) END AS [EnvelopeNumber]
, eo.Description AS [EnvelopeChoice]
, fp.Description AS [FamilyPosition]
From People p
LEFT JOIN People sp on sp.PeopleId = p.SpouseId
INNER JOIN Families fam1 ON p.FamilyId = fam1.FamilyId
LEFT JOIN EnvNum en ON en.PeopleId = p.PeopleId
LEFT JOIN lookup.EnvelopeOption eo ON eo.Id = p.EnvelopeOptionsId
LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
Where (p.EnvelopeOptionsId = 2 or sp.EnvelopeOptionsId = 2)
AND (fam1.HeadOfHouseholdId = (p.PeopleId))
AND ((NOT (p.IsDeceased = 1)) AND (NOT (sp.IsDeceased = 1)))
AND (NOT (p.ArchivedFlag = 1))
AND p.PrimaryAddress <> ''
AND (NOT (p.DoNotMailFlag = 1))
),

EnvJointDeceased AS (
Select 
  CONCAT(p.TitleCode, ' ', p.FirstName, ' ', p.LastName) AS [Name]
, p.PrimaryAddress
, p.PrimaryAddress2
, p.PrimaryCity
, p.PrimaryState
, p.PrimaryZip
, p.SpouseId
, p.PeopleId
, CASE WHEN en.IntValue is NULL THEN cast(p.PeopleId as varchar) ELSE CONCAT(cast(en.IntValue AS varchar), '-', cast(p.PeopleId AS varchar)) END AS [EnvelopeNumber]
, eo.Description AS [EnvelopeChoice]
, fp.Description AS [FamilyPosition]
From People p
LEFT JOIN People sp on sp.PeopleId = p.SpouseId
INNER JOIN Families fam1 ON p.FamilyId = fam1.FamilyId
LEFT JOIN EnvNum en ON en.PeopleId = p.PeopleId
LEFT JOIN lookup.EnvelopeOption eo ON eo.Id = p.EnvelopeOptionsId
LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
Where (p.EnvelopeOptionsId = 2 or sp.EnvelopeOptionsId = 2)
AND (fam1.HeadOfHouseholdId = (p.PeopleId))
AND ((p.IsDeceased = 1 OR sp.IsDeceased = 1) OR p.SpouseId IS NULL)
AND (NOT (p.ArchivedFlag = 1))
AND p.PrimaryAddress <> ''
AND (NOT (p.DoNotMailFlag = 1))
),

EnvIndividual AS (
Select 
  CONCAT(p.TitleCode, ' ', p.FirstName, ' ', p.LastName) AS [Name]
, p.PrimaryAddress
, p.PrimaryAddress2
, p.PrimaryCity
, p.PrimaryState
, p.PrimaryZip
, CASE WHEN en.IntValue is NULL THEN cast(p.PeopleId as varchar) ELSE CONCAT(cast(en.IntValue AS varchar), '-', cast(p.PeopleId AS varchar)) END AS [EnvelopeNumber]
, eo.Description AS [EnvelopeChoice]
, fp.Description AS [FamilyPosition]
From People p
LEFT JOIN People sp on sp.PeopleId = p.SpouseId
INNER JOIN Families fam1 ON p.FamilyId = fam1.FamilyId
LEFT JOIN EnvNum en ON en.PeopleId = p.PeopleId
LEFT JOIN lookup.EnvelopeOption eo ON eo.Id = p.EnvelopeOptionsId
LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
Where (p.EnvelopeOptionsId = 1)
AND (NOT (p.IsDeceased = 1))
AND (NOT (p.ArchivedFlag = 1))
AND p.PrimaryAddress <> ''
AND (NOT (p.DoNotMailFlag = 1))
),

EnvChild AS (
Select 
  CONCAT(p.FirstName, ' ', p.LastName) AS [Name]--, ' and ', s.TitleCode, ' ', s.FirstName, ' ', p.LastName) AS [Name]
, p.PrimaryAddress
, p.PrimaryAddress2
, p.PrimaryCity
, p.PrimaryState
, p.PrimaryZip
, p.Age
, CASE WHEN en.IntValue is NULL THEN cast(p.PeopleId as varchar) ELSE CONCAT(cast(en.IntValue AS varchar), '-', cast(p.PeopleId AS varchar)) END AS [EnvelopeNumber]
, eo.Description AS [EnvelopeChoice]
, fp.Description AS [FamilyPosition]
FROM People p 
LEFT JOIN EnvNum en ON en.PeopleId = p.PeopleId
LEFT JOIN lookup.EnvelopeOption eo ON eo.Id = p.EnvelopeOptionsId
LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
WHERE p.EnvelopeOptionsId IN (1,2) and p.Age < 18 and p.Age is not null)

Select 
ej.Name
,ej.PrimaryAddress
,ej.PrimaryAddress2
,ej.PrimaryCity
,ej.PrimaryState
,ej.PrimaryZip
,ej.EnvelopeNumber
,ej.EnvelopeChoice
,ej.FamilyPosition
FROM
EnvJoint ej
UNION ALL
Select 
ejd.Name
,ejd.PrimaryAddress
,ejd.PrimaryAddress2
,ejd.PrimaryCity
,ejd.PrimaryState
,ejd.PrimaryZip
,ejd.EnvelopeNumber
,ejd.EnvelopeChoice
,ejd.FamilyPosition
FROM
EnvJointDeceased ejd
UNION ALL
Select 
ei.Name
,ei.PrimaryAddress
,ei.PrimaryAddress2
,ei.PrimaryCity
,ei.PrimaryState
,ei.PrimaryZip
,ei.EnvelopeNumber
,ei.EnvelopeChoice
,ei.FamilyPosition
FROM
EnvIndividual ei
