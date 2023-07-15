ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

model.Header = 'Table Filter Test'

listsql = '''
SELECT 
    dbo.People.FamilyId, 
    dbo.People.PeopleId, 
    dbo.People.Name AS [MemberName], 
    dbo.TransactionSummary.TranDate, 
    Organizations_alias1.OrganizationName, 
    dbo.TransactionSummary.TotDue,
    Organizations_alias1.OrganizationId
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
    dbo.TransactionSummary 
ON 
    ( 
        Organizations_alias1.OrganizationId = dbo.TransactionSummary.OrganizationId) 
LEFT OUTER JOIN 
    dbo.People 
ON 
    ( 
        dbo.TransactionSummary.PeopleId = dbo.People.PeopleId) 
WHERE 
    dbo.Program.Id = ''' + ProgramID + ''' 
Order By dbo.People.LastName, dbo.People.FirstName;
'''

print '''

<script src ="https://code.jquery.com/jquery-3.5.1.min.js"></script>                 
<script src ="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"> </script>            
<style>    
h2 {    
  margin-top: 40px;    
  text-transform: none;    
  font-size: 1.75em;    
  font-weight: bold;    
   font-family: "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif;    
  color: white;     
  letter-spacing: -0.005em;     
  word-spacing: 1px;    
  letter-spacing: none;    
  text-align: center;    
}    
h4 {    
  margin-top: 40px;    
  text-transform: none;    
  font-size: 1.75em;    
  font-weight: bold;    
   font-family: "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif;    
  color: #999;     
  letter-spacing: -0.005em;     
  word-spacing: 1px;    
  letter-spacing: none;    
  text-align: center;    
  color: white;    
}    
body {  
  background: #ddd;  
  font-family: sans-serif;  
}  
.container {  
  margin: 20px 0 20px 0;  
}  
.table-filter {  
  margin-bottom: 5px;  
}  
table {  
  box-shadow: 0px 0px 20px rgba(0, 0, 0, 0.1), 0px 10px 20px rgba(0, 0, 0, 0.05),  
    0px 20px 20px rgba(0, 0, 0, 0.05), 0px 30px 20px rgba(0, 0, 0, 0.05);  
}  
th {  
  color: #ffffff;  
  background: #39a2fb;  
  font-weight: 700;  
}  
tr {  
  background: #fff;  
}  
tr:hover {  
  background: #f4f4f4;  
}  
th {  
  padding: 5px 5px 5px 10px;  
}  
td {  
  padding: 5px 5px 5px 10px;  
}  
td  {  
  word-wrap: break-word;  
  border-bottom: 1px solid #ccc;  
}  
</style>  
<body>  
<section class = "container">  
<h2> Example </h2>  
<h4> Jquery filter table </h4>  
  <input  type = "search"  
    class = "form-control table-filter"  
    placeholder = "Search..."  
  />  
  <div class = "table-responsive">  
    <table role = "table" class = "table filtered-table">  
      <thead role = "rowgroup">  
        <tr role = "row">  
          <th role = "columnheader"> FamilyId </th>  
          <th role = "columnheader"> PeopleId </th>  
          <th role = "columnheader"> Name </th>  
          <th role = "columnheader"> Fee Group </th>  
          <th role = "columnheader"> Total Due </th>  
          <th role = "columnheader">  Paylink </th>  
        </tr>  
      </thead>  
      <tbody role = "rowgroup">  

'''

for a in q.QuerySql(listsql):
    
    due = a.TotDue
    if due != 0:
        paylink = model.GetPayLink(a.PeopleId, a.OrganizationId)
        paylink = model.GetAuthenticatedUrl(a.PeopleId, paylink, True)
    
    #if paylink != None:
    print '<tr role = "row">'
    print '<td role = "cell"> {} </td>'.format(a.FamilyId)
    print '<td role = "cell"> {} </td>'.format(a.PeopleId) 
    print '<td role = "cell"> {} </td>'.format(a.MemberName) 
    print '<td role = "cell"> {} </td>'.format(a.OrganizationName) 
    print '<td role = "cell"> ${} </td>'.format(a.TotDue) 
    if due != 0:
        print '<td role = "cell">'
        print '<a href="#" onclick="prepText({0}, this);">Text</a>'.format(paylink) 
        print '<a href="#" onclick="prepText({0}, this);"> Email</a>'.format(paylink) 
        print '<a href="{0}" target="_private"> Self</a>'.format(paylink) 
        print '</td>'
    
    else:
        print '<td role = "cell"></td>'
        
    print "<tr>"
    #print "<h5>{} | {} | Due: ${}</h3>".format(a.MemberName, a.OrganizationName, a.TotDue)
    #print '''<a href="#" onclick="prepText({0}, '{1}', {2}, this);" class="btn btn-primary">Text Paylink</a>'''.format(a.PeopleId, paylink, a.TotDue)
    #print "</div><br>"
    #print '<div class="form-group">'
    #print '<label for="searchname" class="control-label">Name</label>'
    #print '''<button type="button" onclick='searchPerson(searchname.innerHTML)'>+</button>'''
    #print '<input type="test" name="refId" id="refId" class="form-control">'


print '''
      </tbody>  
    </table>  
  </div>  
</section>  
<script>
(function ($) {  
  'use strict';  
  $(document).ready(function () {  
    $('.table-filter').on('input', function () {  
      var value = $(this).val().toLowerCase();  
      $('.filtered-table tbody tr').filter(function () {  
        $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1);  
      });  
    });  
  });  
})(window.jQuery);  
</script>  
'''

#Data.list = q.QuerySql(listsql)