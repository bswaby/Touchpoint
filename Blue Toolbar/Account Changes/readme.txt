Account changes is meant to be a blue toolbar tool to allow you to run account changes against a person, involvement, or search.  

Copy AccountChanges.sql to a new SQL script under Admin ~ Special Content ~ Sql Scripts and call it AccountChanges

Open CustomReport under Admin ~ Special Content ~ Text Content and past in the following line.  Be sure it's under <CustomReports>
  <Report name="AccountChanges" type="SqlReport" role="Admin" />

For role add in the Roles that should see this seperated by a comma.

Note:  Due to how TP implemented replication of CustomReports changes, it could take up to 24hrs before the menu item show.



