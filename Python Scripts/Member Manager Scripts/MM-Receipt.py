model.Title = "Receipt"

pid = model.Data.p1
tranid = model.Data.TranId
ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

transactions = '''
    Select 
            t.Id,
            t.TransactionDate,
            t.TransactionId,
            REPLACE(
                REPLACE(
                REPLACE(t.message, 'APPROVED', 'Online Transaction')
                       , 'Response: ',  '')
                       ,'CC -', 'CC |') as Description,
            FORMAT(t.amt, 'C') AS [amt],
            t.AdjustFee,
            ts.PeopleId,
            ts.OrganizationId,
            org.OrganizationName,
            ts.RegId,
            p.FirstName,
            p.LastName,
            p.EmailAddress
        From TransactionSummary ts
        Left Join [Transaction] t
        ON t.originalId = ts.regid 
        Left Join Organizations org
        On ts.OrganizationId = org.OrganizationId
        INNER Join People p
        ON ts.PeopleId = p.PeopleId
        where 
          t.Id = {1}
        Order by TransactionDate
'''.format(pid,tranid)

transactionsnew = '''
SELECT 
    t.Id, 
    t.TransactionDate, 
    t.TransactionId,         
    REPLACE(
            REPLACE(
            REPLACE(t.message, 'APPROVED', 'Online Transaction')
                   , 'Response: ',  '')
                   ,'CC -', 'CC |') as Description,
    FORMAT(t.amt, 'C') AS [amt],
    t.AdjustFee,
    tp.PeopleId,
    tp.OrgId,
    org.OrganizationName,
    p.FirstName,
    p.LastName,
    p.EmailAddress
    --tp.RegId 
FROM [Transaction] t
LEFT JOIN [TransactionPeople] tp ON t.OriginalId = tp.Id
LEFT JOIN Organizations org On t.OrgId = org.OrganizationId
LEFT JOIN People p ON p.PeopleId = tp.PeopleId
WHERE t.Id = {1} --(tp.PeopleId = {1}) AND (t.AdjustFee = 0 OR t.TransactionGateway <> ' ') AND t.amt <> 0 
'''.format(pid,tranid)

for a in q.QuerySql(transactionsnew):
    message = '''
        <div class="bvedit">
        <!--[if gte mso 9]>
        <xml>
          <o:OfficeDocumentSettings>
            <o:AllowPNG></o:AllowPNG>
            <o:PixelsPerInch>96</o:PixelsPerInch>
          </o:OfficeDocumentSettings>
        </xml>
        <![endif]-->
          <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <meta name="x-apple-disable-message-reformatting">
          <!--[if !mso]><!--><meta http-equiv="X-UA-Compatible" content="IE=edge"><!--<![endif]-->
          <title></title>
          
            <style type="text/css">
              @media only screen and (min-width: 520px) {
          .u-row {
            width: 500px !important;
          }
          .u-row .u-col {
            vertical-align: top;
          }
        
          .u-row .u-col-100 {
            width: 500px !important;
          }
        
        }
        
        @media (max-width: 520px) {
          .u-row-container {
            max-width: 100% !important;
            padding-left: 0px !important;
            padding-right: 0px !important;
          }
          .u-row .u-col {
            min-width: 320px !important;
            max-width: 100% !important;
            display: block !important;
          }
          .u-row {
            width: calc(100% - 40px) !important;
          }
          .u-col {
            width: 100% !important;
          }
          .u-col > div {
            margin: 0 auto;
          }
        }
        body {
          margin: 0;
          padding: 0;
        }
        
        table,
        tr,
        td {
          vertical-align: top;
          border-collapse: collapse;
        }
        
        p {
          margin: 0;
        }
        
        .ie-container table,
        .mso-container table {
          table-layout: fixed;
        }
        
        * {
          line-height: inherit;
        }
        
        a[x-apple-data-detectors='true'] {
          color: inherit !important;
          text-decoration: none !important;
        }
        
        table, td { color: #000000; } </style>
        
          <!--[if IE]><div class="ie-container"><![endif]-->
          <!--[if mso]><div class="mso-container"><![endif]-->
          <table style="border-collapse: collapse;table-layout: fixed;border-spacing: 0;mso-table-lspace: 0pt;mso-table-rspace: 0pt;vertical-align: top;min-width: 320px;Margin: 0 auto;background-color: #e7e7e7;width:100%" cellpadding="0" cellspacing="0">
          <tbody>
          <tr style="vertical-align: top">
            <td style="word-break: break-word;border-collapse: collapse !important;vertical-align: top">
            <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" style="background-color: #e7e7e7;"><![endif]-->
            
        
        <div class="u-row-container" style="padding: 0px;background-color: #ced4d9">
          <div class="u-row" style="Margin: 0 auto;min-width: 320px;max-width: 500px;overflow-wrap: break-word;word-wrap: break-word;word-break: break-word;background-color: #ced4d9;">
            <div style="border-collapse: collapse;display: table;width: 100%;height: 100%;background-color: transparent;">
              <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding: 0px;background-color: #ced4d9;" align="center"><table cellpadding="0" cellspacing="0" border="0" style="width:500px;"><tr style="background-color: #ced4d9;"><![endif]-->
              
        <!--[if (mso)|(IE)]><td align="center" width="500" style="width: 500px;padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;" valign="top"><![endif]-->
        <div class="u-col u-col-100" style="max-width: 320px;min-width: 500px;display: table-cell;vertical-align: top;">
          <div style="height: 100%;width: 100% !important;">
          <!--[if (!mso)&(!IE)]><!--><div style="height: 100%; padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;"><!--<![endif]-->
          
        <table style="font-family:arial,helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
          <tbody>
            <tr>
              <td style="overflow-wrap:break-word;word-break:break-word;padding:10px;font-family:arial,helvetica,sans-serif;" align="left">
                
          <div style="line-height: 140%; text-align: left; word-wrap: break-word;">
            
          </div>
        
              </td>
            </tr>
          </tbody>
        </table>
        
          <!--[if (!mso)&(!IE)]><!--></div><!--<![endif]-->
          </div>
        </div>
        <!--[if (mso)|(IE)]></td><![endif]-->
              <!--[if (mso)|(IE)]></tr></table></td></tr></table><![endif]-->
            </div>
          </div>
        </div>
  
        <div class="u-row-container" style="padding: 0px;background-color: #ced4d9">
          <div class="u-row" style="Margin: 0 auto;min-width: 320px;max-width: 500px;overflow-wrap: break-word;word-wrap: break-word;word-break: break-word;background-color: #ffffff;">
            <div style="border-collapse: collapse;display: table;width: 100%;height: 100%;background-color: transparent;">
              <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding: 0px;background-color: #ced4d9;" align="center"><table cellpadding="0" cellspacing="0" border="0" style="width:500px;"><tr style="background-color: #ffffff;"><![endif]-->
              
        <!--[if (mso)|(IE)]><td align="center" width="500" style="width: 500px;padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;" valign="top"><![endif]-->
        <div class="u-col u-col-100" style="max-width: 320px;min-width: 500px;display: table-cell;vertical-align: top;">
          <div style="height: 100%;width: 100% !important;">
          <!--[if (!mso)&(!IE)]><!--><div style="height: 100%; padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;"><!--<![endif]-->
          
        <table style="font-family:arial,helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
          <tbody>
            <tr>
              <td style="overflow-wrap:break-word;word-break:break-word;padding:10px;font-family:arial,helvetica,sans-serif;" align="left">
                
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tbody><tr>
            <td style="padding-right: 0px;padding-left: 0px;" align="center">
              
              <img align="center" border="0" src="https://fbchville.com/wp-content/uploads/2022/02/MenuLogo.png" alt="Image" title="Image" style="outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;clear: both;display: inline-block !important;border: none;height: auto;float: none;width: 30%;max-width: 144px;" width="144">
              
            </td>
          </tr>
        </tbody></table>
        
              </td>
            </tr>
          </tbody>
        </table>
        
          <!--[if (!mso)&(!IE)]><!--></div><!--<![endif]-->
          </div>
        </div>
        <!--[if (mso)|(IE)]></td><![endif]-->
              <!--[if (mso)|(IE)]></tr></table></td></tr></table><![endif]-->
            </div>
          </div>
        </div>
        
        
        
        <div class="u-row-container" style="padding: 0px;background-color: #ced4d9">
          <div class="u-row" style="Margin: 0 auto;min-width: 320px;max-width: 500px;overflow-wrap: break-word;word-wrap: break-word;word-break: break-word;background-color: #ffffff;">
            <div style="border-collapse: collapse;display: table;width: 100%;height: 100%;background-color: transparent;">
              <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding: 0px;background-color: #ced4d9;" align="center"><table cellpadding="0" cellspacing="0" border="0" style="width:500px;"><tr style="background-color: #ffffff;"><![endif]-->
              
        <!--[if (mso)|(IE)]><td align="center" width="500" style="width: 500px;padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;" valign="top"><![endif]-->
        <div class="u-col u-col-100" style="max-width: 320px;min-width: 500px;display: table-cell;vertical-align: top;">
          <div style="height: 100%;width: 100% !important;">
          <!--[if (!mso)&(!IE)]><!--><div style="height: 100%; padding: 0px;border-top: 0px solid transparent;border-left: 0px solid transparent;border-right: 0px solid transparent;border-bottom: 0px solid transparent;"><!--<![endif]-->
          
        <table style="font-family:arial,helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
          <tbody>
            <tr>
              <td style="overflow-wrap:break-word;word-break:break-word;padding:10px;font-family:arial,helvetica,sans-serif;" align="left">
                
          <table height="0px" align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;table-layout: fixed;border-spacing: 0;mso-table-lspace: 0pt;mso-table-rspace: 0pt;vertical-align: top;border-top: 1px solid #BBBBBB;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%">
            <tbody>
              <tr style="vertical-align: top">
                <td style="word-break: break-word;border-collapse: collapse !important;vertical-align: top;font-size: 0px;line-height: 0px;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%">
                  <span>&nbsp;</span>
                </td>
              </tr>
            </tbody>
          </table>
        
              </td>
            </tr>
          </tbody>
        </table>
        
        <table style="font-family:arial,helvetica,sans-serif;" role="presentation" cellpadding="0" cellspacing="0" width="100%" border="0">
          <tbody>
            <tr>
              <td style="overflow-wrap:break-word;word-break:break-word;padding:10px;font-family:arial,helvetica,sans-serif;" align="left">
                
          <div style="line-height: 140%; text-align: left; word-wrap: break-word;">
            <p style="font-size: 14px; line-height: 140%;">Hi ''' + a.FirstName + ''',</p>
        <p style="font-size: 14px; line-height: 140%;">&nbsp;</p>
        <p style="font-size: 14px; line-height: 140%;">Here is your payment confirmation for transaction on ''' + str(a.TransactionDate) + '''.</p>
        <p style="font-size: 14px; line-height: 140%;">&nbsp;</p>
        <p style="font-size: 14px; line-height: 140%;"><table>
            <tbody>
                <tr>
                    <td colspan="2" style="font-family:Arial,Helvetica;font-size:15px;font-weight:bold;border-bottom: solid 1px #D8D8D8">
                        <p>
                            <hr/>
                             ''' + a.OrganizationName + '''
                        </p>
                    </td>
                </tr>
                <tr>
                    <td style="font-family:Arial,Helvetica;font-size:15px;color:#656565;">Payment</td>
                </tr>
                <tr>
                    <td style="font-family:Arial,Helvetica;font-size:15px;font-weight:bold;"> ''' + a.amt + '''</td>
                </tr>
            </tbody>
        </table>
        </p>
        <p style="font-size: 14px; line-height: 140%;">&nbsp;</p>
        <p style="font-size: 14px; line-height: 140%;">Your Ministry Team</p>
        <p style="font-size: 14px; line-height: 140%;">First Baptist Church Hendersonville</br>
           106 Bluegrass Commons Blvd., Hendersonvillle, TN 37075</br>
           (615) 824-6154</br>
           <a href="https://fbchville.com">https://fbchville.com</a>
        </p>
          </div>
        
              </td>
            </tr>
          </tbody>
        </table>
        
        '''
    #print message
    message = message + '''</br></br></br><link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">'''    
    message = message + '''<input type="button" value=" < " onclick="history.back()">'''
    message = message + '''<button onclick="window.location.href=' ''' + model.CmsHost + '''/PyScript/MM-MemberManager?ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''';"><i class="fa fa-home"></i></button>'''
    message = message + '''<button onclick="window.location.href=' ''' + model.CmsHost + '''/PyScript/MM-ReceiptEmail?P1=''' + str(a.PeopleId) + '''&TranId=''' + str(a.Id) + '''&ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + ''' ';"><i class="fa fa-envelope-o"></i></button>'''
    print message