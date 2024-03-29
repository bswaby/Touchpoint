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
                    <head>
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
              <!--[if !mso]><!-->
              <meta http-equiv="X-UA-Compatible" content="IE=edge">
              <!--<![endif]-->
              <title>FBCH Payment Receipt</title>
              <link rel="icon" type="image/x-icon" href="https://fbchville.com/wp-content/uploads/2022/02/MenuLogo.png">

              <!-- <style type="text/css">
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
                
                        table, td { color: #000000; } 
                    </style> -->

              <style>
                body {
                  background-color: #ccd6d9;
                  font-family: Arial, Helvetica, sans-serif;
                }

                .content-container {
                  display: flex;
                  height: 88vh;
                  margin: 3%;
                }

                .main-column {
                  background-color: white;
                  flex-grow: 3;
                  box-shadow: 5px 5px 20px gray;
                  max-width: 600px;
                  height: 820px

                }

                .side-column {
                  flex-grow: 2;
                }

                .column-header {
                  width: 100%;
                  /* height: 300px; */
                }

                .column-header-content {
                  background-color: white;
                  height: 20%;
                  margin: auto;
                  display: block;
                }

                .column-header-shape {
                  clip-path: polygon(0% 0%, 100% 0%, 100% 32%, 0% 100%);
                  background-color: #579cac;
                  height: 100%;
                }

                .content-header-image {
                  position: relative;
                  top: -130px;
                  width: fit-content;
                  margin: auto;
                  margin-top: 3%;
                  padding: 5px;
                  background-color: white;
                  clip-path: circle(50%);
                }

                .column-content {
                  margin: 2%;
                }

                .receipt-amount {
                  display: flex;
                  flex-direction: row;
                }

                .receipt-amount-middle {
                  flex-grow: 1;
                }

                .column-footer {
                  background-color: #579cac;
                  position: relative;
                  bottom: -215;
                  height: 70px;
                  display:flex;
                  justify-content: center;
                  align-items: center;
                }

                .column-footer-content{
                  width: 100%;
                  height: 100%;
                  border: 1px solid blue;
                  text-align: center;
                }
              </style>

            </head>

            <body>

              <div class="content-container">
                <div class="side-column"></div>
                <div class="main-column">
                  
                  <div class="column-header">
                    <div class="column-header-content">
                      <div class="column-header-shape">
                      </div>
                      <div class="content-header-image">
                        <img height="75" width="75"
                          src="https://fbchville.com/wp-content/uploads/2022/02/cropped-Favicon-1-192x192.png">
                      </div>
                    </div>
                  </div>


                  <div class="column-content">
                    <h1>Thanks for your payment!</h1>
                    <p>Hi ''' + a.FirstName + ''', </p>
                    <p>Here is your payment confirmation for the transaction on ''' + str(a.TransactionDate) + '''.</p>

                    <hr>

                    <div class="receipt-amount">
                      <p>Total:</p>
                      <div class="receipt-amount-middle"></div>
                      <p> $''' + a.AdjustFee + ''' </p>
                    </div>

                    <hr>
                    <p>Your Ministry Team,</p>
                    <p>First Baptist Church Hendersonville</br>
                      106 Bluegrass Commons Blvd., Hendersonvillle, TN 37075</br>
                      (615) 824-6154</br>
                      <a href="https://fbchville.com">https://fbchville.com </a>
                    </p>
                  </div>

                  <!-- <div class="column-footer">
                    <p>First Baptist Hendersonville</p>
                  </div> -->
                  <div class="column-footer">
                    &nbsp;
                  </div>
                  
                </div>
                <div class="side-column">
                </div>



              </div>




            </body>
        '''
    #print message
    message = message + '''</br></br></br><link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">'''    
    message = message + '''<input type="button" value=" < " onclick="history.back()">'''
    message = message + '''<button onclick="window.location.href=' ''' + model.CmsHost + '''/PyScript/MM-MemberManager?ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''';"><i class="fa fa-home"></i></button>'''
    message = message + '''<button onclick="window.location.href=' ''' + model.CmsHost + '''/PyScript/MM-ReceiptEmail?P1=''' + str(a.PeopleId) + '''&TranId=''' + str(a.Id) + '''&ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + ''' ';"><i class="fa fa-envelope-o"></i></button>'''
    print message