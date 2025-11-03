#roles=Finance
#class=UploadContributionsMenu,title=Upload Pushpay

from datetime import datetime

def ProcessGet():
    #form = model.DynamicData()
    btsql = '''
        select Id, Description from lookup.BundleHeaderTypes where Id not in (4,5,6)
    '''
    Data.bundletypes = q.QuerySql(btsql)
    
    fsql = '''
        select FundId [Id], FundName [Name]
        from dbo.ContributionFund where FundStatusId = 1
        {0}
    '''
    
    sort = model.Setting('SortContributionFundsByFieldName')
    if sort == 'FundName':
        order = 'order by FundName'
    else:
        order = 'order by FundId'

    fosql = fsql.format(order)
    Data.funds = q.QuerySql(fosql)

    html = model.Content('UploadContributions')
    model.Form = model.RenderTemplate(html)
    
def GetContTypeId(code):
    contTypeDict = {'c': 1, 'n': 9, 'p': 8, 's': 20, 'g': 10}
    return contTypeDict.get(code, 1) # defaut to ContributionTypeId = 1 if code not found
    
def FindPersonIdFromEnvelopeNumber(value):
    sql = '''
    select top 1
    	PeopleId
    from dbo.PeopleExtra 
    where Field = 'EnvelopeNumber'
    and IntValue = {0}
    '''
    return q.QuerySqlInt(sql.format(value))

def ProcessPost():
    csv = model.CsvReader(model.Data.file)
    dt = model.ParseDate(model.Data.date)
    ddt = model.ParseDate(model.Data.ddate)
    bundleType = int(model.Data.bundleType)
    referenceId = model.Data.refId[:100]
    if bundleType == 0:
        bundleType = 7
    bundleHeader = model.GetBundleHeader(dt, model.DateTime.Now, bundleTypeId = bundleType)
    bundleHeader.DepositDate = ddt
    fid = int(model.Data.fundid or model.FirstFundId())
    if referenceId is not None:
        bundleHeader.ReferenceId = referenceId
    while csv.Read():
        Date = model.ParseDate(csv['Received On'])
        EnvelopeNumber = int(csv['Your ID'])
        PushPayKey = csv['Contributor ID']
        fid1 = csv['Fund Code']
        Amount = csv['Amount']
        CheckNumber = csv['Payment Type']
        TaxDed = csv['Tax Deductible']
        Memo = csv['Memo']
        FirstName = csv['First Name']
        LastName = csv['Last Name']
        Email = csv['Email']
        Cell = csv['Mobile Number'][-10:]
        TransactionId = csv['Transaction ID']
        
        if fid1 is None or fid1 == '':
            fid1 = fid
            
        bd = model.AddContributionDetail(Date, int(fid1), Amount, CheckNumber, None, None)
        
        if Memo is not None and Memo != '':
            bd.Contribution.ContributionDesc = Memo[:256]
        
        if TaxDed.upper() == 'FALSE':
            bd.Contribution.ContributionTypeId = 9
        else:
            bd.Contribution.ContributionTypeId = 1

        pid = model.FindPersonIdExtraValue('PushPayKey', PushPayKey)
        if pid is None:
            pid = FindPersonIdFromEnvelopeNumber(EnvelopeNumber)
            if pid is not None and pid != 0:
                model.AddExtraValueText(pid, 'PushPayKey', PushPayKey)
        if pid is None or pid == 0:
            pid = model.FindPersonId(FirstName, LastName, None, Email, Cell)
            if pid is not None:
                model.AddExtraValueText(pid, 'PushPayKey', PushPayKey)
        bd.Contribution.PeopleId = pid
        
        bd.Contribution.MetaInfo = ("Pushpay Tran #" + TransactionId)[:100]
        bundleHeader.BundleDetails.Add(bd)

    model.FinishBundle(bundleHeader)
    id = bundleHeader.BundleHeaderId
    print('REDIRECT=/Batches/Detail/{0}'.format(id))

if model.HttpMethod == 'get':
    ProcessGet()
elif model.HttpMethod == 'post':
    ProcessPost()
