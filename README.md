<h2>Import Zaereen Data from the Excel Sheet into Ms Access Database.</h2>
<ul><li>Create seperate excel sheet and place the data in it, it should be in the below format.</li>
<ul><li>ITS</li><li>Name</li><li>Mobile</li><li>Age</li><li>Location</li><li>Occupation</li><li>TripExp</li><li>Remarks</li></ul>. <p>In case there are extra fields, then the code should be amended before executing.</p></li>
<li>It should start from the first column i.e.: A1</li>
<li>First row should be the column names and should match the above names.</li>
<li>2 arguments should be passed to the command line (use command prompt to execute the application, enter the path of the application and pass the 2 parameters).
i.e.: Excel Sheet (the one created above) path and the database (create a copy of production database on local) path
both the paths should have double quotes around them and both the paths should be separated by space.</li>
<li>Before attempting to run ensure the Account_No field in the production database is auto-generated, if not, then amend the code to generate account_no dynamically.</li>
<li>Run the application and it will insert the records in the database.</li></ul>
