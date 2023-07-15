select
	p.PeopleId,
	p.FamilyId,
	P.PreferredName,
	p.LastName,
	p.BDate DOB,
	case when p.Age = 17 
	        and datediff(day, getdate(), dbo.NextBirthday(p.PeopleId)) <= 15 then 'True' 
	        else NULL end [Turning 18],
	case when p.Age <= 11 then 'True' else NULL end [11 and Under],
	case when p.Age >= 80 then 'True' else NULL end [80+],
	case sf.F904 when 'x' then 'True' else NULL end [Special Needs],
	case when (select count(*) from dbo.People p2 where p2.FamilyId = p.FamilyId
								and p2.PeopleId in (select PeopleId from dbo.OrganizationMembers
														where OrganizationId = 852)) >= 1 
		then 'True' else NULL end [Staff],
	dn.Data [Discount Notes],
	convert(date,(select max(et.TransactionDate) from dbo.EnrollmentTransaction et 
									where et.OrganizationId = 834
									and et.PeopleId = p.PeopleId
									and et.TransactionTypeId = 1)) [Join Date],
	isnull(ts.IndAmt, 0) [Fee],
	isnull(ts.IndPaid, 0) [Paid],
	isnull(ts.IndDue, 0) [Due]
from dbo.People p
join dbo.TagPerson tp on tp.PeopleId = p.PeopleId and tp.Id = @qtagid
join dbo.OrganizationMembers om on om.PeopleId = p.PeopleId and om.OrganizationId = @CurrentOrgId
join dbo.StatusFlagColumns sf on sf.PeopleId = p.PeopleId
left join dbo.TransactionSummary ts on ts.RegId = om.TranId
left join dbo.PeopleExtra dn on dn.PeopleId = p.PeopleId and dn.Field = (@ProgramName + '-DiscountNotes')