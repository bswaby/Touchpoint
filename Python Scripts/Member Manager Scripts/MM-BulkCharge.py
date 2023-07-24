ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName


#page styling
print '''
<head>
	<style>
		button {
            width: 100%;
			color: #ffffff;
			background-color: #2d63c8;
			font-size: 19px;
			border: 1px solid #2d63c8;
			padding: 15px 50px;
			cursor: pointer
		}
		button:hover {
			color: #2d63c8;
			background-color: #ffffff;
		}
	</style>
</head>
'''

print '''
  <div class = "table-responsive">  
    <table role = "table" class = "table filtered-table">  
      <thead role = "rowgroup">  
        <tr role = "row">  
          <th role = "columnheader"> Name </th>  
          <th role = "columnheader"> Involvement (subgroup) </th>
          <th role = "columnheader"> Outstanding </th>
          <th role = "columnheader"> New Charges </th>
        </tr>  
      </thead>  
      <tbody role = "rowgroup">  

'''