#roles=Finance
#class=ReportsMenu,title=Giver By Group

from datetime import datetime

def ProcessGet():
    #form = model.DynamicData()
    btsql = '''
        SELECT 
            DISTINCT Organizations_alias1.OrganizationName, 
            Organizations_alias1.OrganizationId 
        FROM 
            dbo.Organizations Organizations_alias1 
        ORDER BY 
            Organizations_alias1.OrganizationName ASC;
    '''
    Data.bundletypes = q.QuerySql(btsql)
    
    fsql = '''
        select FundId [Id], FundName [Name]
        from dbo.ContributionFund where FundStatusId = 1
        {0}
    '''
    
    #sort = model.Setting('SortContributionFundsByFieldName')
    #if sort == 'FundName':
    #    order = 'order by FundName'
    #else:
    #    order = 'order by FundId'

    #fosql = fsql.format(order)
    #Data.funds = q.QuerySql(fosql)

    html = model.Content('GivingGroups')
    model.Form = model.RenderTemplate(html)
    
def GetContTypeId(code):
    contTypeDict = {'c': 1, 'n': 9, 'p': 8, 's': 20, 'g': 10}
    return contTypeDict.get(code, 1) # defaut to ContributionTypeId = 1 if code not found
    
def ProcessPost():
    #csv = model.CsvReader(model.Data.file)
    #dt = model.ParseDate(model.Data.date)
    #ddt = model.ParseDate(model.Data.ddate)
    OrgID = (model.Data.bundleType)
    #referenceId = model.Data.refId[:100]
    #if bundleType == 0:
    #    bundleType = 4
    #bundleHeader = model.GetBundleHeader(dt, model.DateTime.Now, bundleTypeId = bundleType)
    #bundleHeader.DepositDate = ddt
    #fid = int(model.Data.fundid or model.FirstFundId())
    #if referenceId is not None:
    #    bundleHeader.ReferenceId = referenceId
    #while csv.Read():
    #    Date = model.ParseDate(csv['Received On'])
    #    EnvelopeNumber = int(csv['Your ID'])
    #    PushPayKey = csv['Contributor ID']
    #    fid1 = csv['Fund Code']
    #    Amount = csv['Amount']
    #    CheckNumber = csv['Payment Type']
    #    TaxDed = csv['Tax Deductible']
    #    Memo = csv['Memo']
    #    FirstName = csv['First Name']
    #    LastName = csv['Last Name']
    #    Email = csv['Email']
    #    Cell = csv['Mobile Number'][-10:]
    #    TransactionId = csv['Transaction ID']
    #    
    #    if fid1 is None or fid1 == '':
    #        fid1 = fid
    #        
    #    bd = model.AddContributionDetail(Date, int(fid1), Amount, CheckNumber, None, None)
    #    
    #    if Memo is not None and Memo != '':
    #        bd.Contribution.ContributionDesc = Memo[:256]
    #    
    #    if TaxDed.upper() == 'FALSE':
    #        bd.Contribution.ContributionTypeId = 9
    #    else:
    #        bd.Contribution.ContributionTypeId = 5

    #    pid = model.FindPersonIdExtraValue('PushPayKey', PushPayKey)
    #    if pid is None:
    #        pid = model.FindPersonIdExtraValueInt('EnvelopeNumber', EnvelopeNumber)
    #        if pid is not None:
    #            model.AddExtraValueText(pid, 'PushPayKey', PushPayKey)
    #    if pid is None:
    #        pid = model.FindPersonId(FirstName, LastName, None, Email, Cell)
    #        if pid is not None:
    #            model.AddExtraValueText(pid, 'PushPayKey', PushPayKey)
    #    bd.Contribution.PeopleId = pid
    #    
    #    bd.Contribution.MetaInfo = ("Pushpay Tran #" + TransactionId)[:100]
    #    bundleHeader.BundleDetails.Add(bd)

    #model.FinishBundle(bundleHeader)
    #id = bundleHeader.BundleHeaderId
    #print('REDIRECT=/PostBundle/{0}'.format(id))
    sqlinvolvementgiver = '''
        SELECT 
            lookup.MemberType.Description as [GroupMemberType],
            Max(dbo.Contribution.ContributionDate) as [LastGiftDate],
            SUM(dbo.Contribution.ContributionAmount) as [ContributionAmount],
            dbo.People.Name,
            lookup.MemberStatus.Description as [MemberStatus]
        FROM 
            dbo.OrganizationMembers 
        INNER JOIN 
            dbo.People 
        ON 
            ( 
                dbo.OrganizationMembers.PeopleId = dbo.People.PeopleId) 
        LEFT OUTER JOIN 
            dbo.Contribution 
        ON 
            ( 
                dbo.People.PeopleId = dbo.Contribution.PeopleId) 
        OR 
            ( 
                dbo.People.SpouseId = dbo.Contribution.PeopleId) 
        INNER JOIN 
            lookup.MemberType 
        ON 
            ( 
                dbo.OrganizationMembers.MemberTypeId = lookup.MemberType.Id) 
        INNER JOIN 
            lookup.MemberStatus 
        ON 
            ( 
                dbo.People.MemberStatusId = lookup.MemberStatus.Id) 
        WHERE 
            dbo.OrganizationMembers.OrganizationId = '''+ OrgID + ''' 
        GROUP BY 
            dbo.People.Name,lookup.MemberType.Description,lookup.MemberStatus.Description
        Order by 
            [GroupMemberType],[LastGiftDate], [MemberStatus];
    '''
    template = '''
    <table width="1000px" border="0" cellpadding="5" style="border:1px solid black; border-collapse:collapse">
        <tbody>
        <tr style="{{Bold}}">
                <td  border: 0px solid #000; vertical-align:top; style="background-color: #eeeeee"><font size=4>Involvement<br>Member Type</font></td>
                <td  border: 0px solid #000; vertical-align:top; style="background-color: #eeeeee"><font size=4>Last Give Date</font></td>
                <td  border: 0px solid #000; vertical-align:top;  align="right" style="background-color: #eeeeee"><font size=4>Contribution<br>Amount</font></td>
                <td  border: 0px solid #000; vertical-align:top; align="right" style="background-color: #eeeeee"><font size=4>Name</font></td>
                <td  border: 0px solid #000; vertical-align:top; align="right" style="background-color: #eeeeee"><font size=4>Member Status</font></td>
        </tr>
        {{#each involvementgiver}}
        <tr style="{{Bold}}">
                <td  border: 0px solid #000; vertical-align:top;><font size=3>{{GroupMemberType}}</font></td>
                <td  border: 0px solid #000; vertical-align:top;><font size=3>{{LastGiftDate}}</font></td>
                <td border: 0px solid #000; vertical-align:top; align="right"><font size=3>{{ContributionAmount}}</font></td>
                <td border: 0px solid #000; vertical-align:top; align="right"><font size=3>{{Name}}</font></td>
                <td border: 0px solid #000; vertical-align:top; align="right"><font size=3>{{MemberStatus}}</font></td>
        </tr>
       {{/each}}
        </tbody>
    </table>
    '''
    #print(sqlinvolvementgiver)
    Data.involvementgiver = q.QuerySql(sqlinvolvementgiver)
    SReport = model.RenderTemplate(template)
    print(SReport)

if model.HttpMethod == 'get':
    ProcessGet()
elif model.HttpMethod == 'post':
    ProcessPost()
