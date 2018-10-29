'********************************************************************************************** 
'**********************************************************************************************
' DBFunctions.vbs
' Functions - See individual Function Headers for 
'             a detail description of function functionality
'       executeDBQuery
'		executeDBUpdate
'       executeExcelQuery
'       executeCSVQuery
'       executeMySQLQuery
'       
'**********************************************************************************************
'**********************************************************************************************
Option Explicit

'<@class_summary>
'**********************************************************************************************
' <@purpose>
'   This Class is used to interact with Application Database(s) - Data Recordset generators
'   and random number generator.
'   Execution of this Class File will create a DBFunctions Object automatically
'   Object Name equals "db"
' </@purpose>
'
' <@author>
'		Mike Millgate
' </@author>
'
' <@creation_date>
'   07-27-2006
' </@creation_date>
'
'**********************************************************************************************
'</@class_summary>
Class DBFunctions

	'<@comments>
	'**********************************************************************************************
	' <@name>
	' executeDBQuery
	' </@name>
	'
	' <@purpose>
	' To create a connection to the database and execute a sql statement
	' </@purpose>
	'
	' <@parameters>
	'    sQuery = (ByVal) String - A SQL statement to query a database
	'    sSchema = (ByVal) String - Name of the schema you want to execute against
	'             Valid Values:
	'               user - UDM_USER_DATA
	'               sys - UDM_SYS_DATA
	' </@parameters>
	'
	' <@return>
	'          An adodb recordset object containing the result 
	'          of the database query.  Since the return item is
	'          an object it must be set instead of assigned.
	'          Returns nothing if function failures
	' </@return>
	'
	' <@assumptions>
	'   Environment Variables have been initialized/created
	'     Environment("CONNECTION_TYPE")
	'     Environment("CONNECTION_PROVIDER")
	'     Environment("XTSN")
	'     Environment("SYS_SCHEMA_NAME")
	'     Environment("SYS_SCHEMA_PSWD")
	' </@assumptions>
	'    
	' <@example_usage>
	'  set oRS = db.executeDBQuery("select * from table", "sys")	
	'
	'	 If LCase(typeName(oRS)) <> "recordset" Then
	'			Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
	'  ElseIf oRS.bof And oRS.eof Then
	'			Reporter.ReportEvent micFail, "invalid recordset", "The returned recordset contains no records."
	'  Else
	'		  Reporter.ReportEvent micPass, "valid recordset", "The returned recordset is valid and contains records."
	'  		While Not oRS.eof
	'		  	sCode = oRS.fields(0).Value
	'				oRS.move(1)
	'	 		Wend
	'  End If
	' </@example_usage>
	'
	' <@author>
	' 	Mike Millgate
	' </@author>
	'
	' <@creation_date>
	'		07-27-2006
	' </@creation_date>
	'
	' <@mod_block>
	' 	02-28-2007 - MM - Added name of function to error messages where invalid parameters are passed
	'   10-22-2007 - MM - Added Transaction Timer Command
	'   10-25-2007 - MM - Added logic to check the connection state before executing query
	'   03-18-2008 - MM - Added ByVal References to the function parameters
	' </@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Function executeDBQuery(ByVal sQuery, ByVal sSchema) ' <@as> Recordset
	
	   Services.StartTransaction "executeDBQuery" ' Timer Begin
	   Reporter.ReportEvent micDone, "executeDBQuery Function", "Function begin"
	   
	   ' Variable Declaration / Initialization
	   Dim oConnection
	   
	   ' Check to verify passed parameters that they are not null or an empty string
	   If IsNull(sQuery) or sQuery = "" or IsNull(sSchema) or sSchema = "" Then
	           Reporter.ReportEvent micFail, "Invalid Parameters", "Invalid parameters were passed to the executeDBQuery function check passed parameters"
	           Set executeDBQuery = Nothing ' Return Value
	           Services.EndTransaction "executeDBQuery" ' Timer End
	           Exit Function
	   End If
	   
	   ' Set schema variable to Upper Case
	   sSchema = UCase(sSchema)
	   
	   ' Create Connection to database
	   Set oConnection = CreateObject(Environment.Value("CONNECTION_TYPE")) ' Connection Object
	   oConnection.Provider = Environment.Value("CONNECTION_PROVIDER") ' Provider Type (i.e. Oracle, MS Access, etc.)
	   ' Opens Database Connection (TNSNAMES, DB Schema Name, DB Schema Password)
	   oConnection.open Environment.Value("XTSN"), Environment.Value(sSchema & "_SCHEMA_NAME"), Environment.Value(sSchema & "_SCHEMA_PSWD")
	   
	 	 ' Execute sql statement, if database connection is open
		 If oConnection.State = 1 Then
		 	Set executeDBQuery = oConnection.execute(sQuery) ' Return Value
		 Else ' Connection not open
	   	Reporter.ReportEvent micFail, "connection state", "The database connection did not open."
	   	Set executeDBQuery = Nothing ' Return Value
	 	 End If
	  
	   Reporter.ReportEvent micDone, "executeDBQuery Function", "Function End"
	   Services.EndTransaction "executeDBQuery" ' Timer End
	   
	End Function
	
	'<@comments>
	'**********************************************************************************************
	' <@name>
	' executeDBUpdate
	' </@name>
	'
	' <@purpose>
	' To create a connection to the database and execute a sql statement
	' </@purpose>
	'
	' <@parameters>
	'    sQuery = (ByVal) String - A SQL statement to query a database
	'    sSchema = (ByVal) String - Name of the schema you want to execute against
	'             Valid Values:
	'               user - UDM_USER_DATA
	'               sys - UDM_SYS_DATA
	' </@parameters>
	'
	' <@return>
	'          Nothing
	' </@return>
	'
	' <@assumptions>
	'   Environment Variables have been initialized/created
	'     Environment("CONNECTION_TYPE")
	'     Environment("CONNECTION_PROVIDER")
	'     Environment("XTSN")
	'     Environment("SYS_SCHEMA_NAME")
	'     Environment("SYS_SCHEMA_PSWD")
	' </@assumptions>
	'    
	' <@example_usage>
	'  Call db.executeDBUpdate("select * from table", "sys")
	'
	' <@author>
	' 	Kevin Webb
	' </@author>
	'
	' <@creation_date>
	'		2-6-2009
	' </@creation_date>
	'
	' <@mod_block>
	'
	' </@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Function executeDBUpdate(ByVal sQuery, ByVal sSchema) ' <@as> Recordset
	
	   Services.StartTransaction "executeDBUpdate" ' Timer Begin
	   Reporter.ReportEvent micDone, "executeDBUpdate Function", "Function begin"
	   
	   ' Variable Declaration / Initialization
	   Dim oConnection
	   
	   ' Check to verify passed parameters that they are not null or an empty string
	   If IsNull(sQuery) or sQuery = "" or IsNull(sSchema) or sSchema = "" Then
			   Reporter.ReportEvent micFail, "Invalid Parameters", "Invalid parameters were passed to the executeDBUpdate function check passed parameters"
			   Set executeDBUpdate = Nothing ' Return Value
			   Services.EndTransaction "executeDBUpdate" ' Timer End
			   Exit Function
	   End If
	   
	   ' Set schema variable to Upper Case
	   sSchema = UCase(sSchema)
	   
	   ' Create Connection to database
	   Set oConnection = CreateObject(Environment.Value("CONNECTION_TYPE")) ' Connection Object
	   oConnection.Provider = Environment.Value("CONNECTION_PROVIDER") ' Provider Type (i.e. Oracle, MS Access, etc.)
	   ' Opens Database Connection (TNSNAMES, DB Schema Name, DB Schema Password)
	   oConnection.open Environment.Value("XTSN"), Environment.Value(sSchema & "_SCHEMA_NAME"), Environment.Value(sSchema & "_SCHEMA_PSWD")
	   
		 ' Execute sql statement, if database connection is open
		 If oConnection.State = 1 Then
			oConnection.execute(sQuery)
		 Else ' Connection not open
			Reporter.ReportEvent micFail, "connection state", "The database connection did not open."
		End If
	  
	   Reporter.ReportEvent micDone, "executeDBUpdate Function", "Function End"
	   Services.EndTransaction "executeDBUpdate" ' Timer End
	   
	End Function
	
	'<@comments>	
	'**********************************************************************************************
	' <@name>
	'	executeExcelQuery
	' </@name>
	'
	' <@purpose>
	'	creates a connection to the excel spreadsheet and executes a sql statement.
	' </@purpose>
	'
	' <@parameters>
	'	sFilePath = (ByVal) string - the path to the excel file.
	'        			sSQL = (ByVal) string -	the SQL statement to query the contents of the excel spreadsheet.
	' </@parameters>
	'
	' <@return>
	'	an adodb recordset object containing the result of the excel spreadsheet query.  
	'						since the return item is an object, it must be set instead of assigned.
	'           Returns nothing if function failures
	' </@return>
	'
	' <@assumptions>
	'	environment variables have been initialized/created
	' </@assumptions>
	'
	' <@example_usage>
	'	Set oResultSet = db.executeExcelQuery("c:\file.xls", "select * from [Sheet1$]")
	'
	'									If LCase(typeName(oResultSet)) <> "recordset" Then
	'										Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
	'
	'									ElseIf oResultSet.bof And oResultSet.eof Then
	'										Reporter.ReportEvent micFail, "invalid recordset", "The returned recordset contains no records."
	'
	'									Else
	'										Reporter.ReportEvent micPass, "valid recordset", "The returned recordset is valid and contains records."
	'										
	'										While Not oResultSet.eof
	'											sCode = oResultSet.fields(0).Value
	'											oResultSet.move(1)
	'										Wend
	'
	'									End If
	' </@example_usage>
	'
	' <@author>
	'	craig cardall
	' </@author>
	'
	' <@creation_date>
	'	7/30/2007
	' </@creation_date>
	'
	' <@mod_block>
	'	   10/17/2007 - CC - added a check to verify that the connection state is set to 1 (open)
	'                      before executing the query.  if the connection is closed, the application
	'                      closes and the test exits.
	'		 10/23/2007 - CC - removed the "logoutClose.vbs" file and the corresponding function call.
	'                      added code to set the return value to "nothing" if invalid parameters
	'                      are passed or the database connection doesn't open.
	'    10-25-2007 - MM - Added Transaction Timer Command
	'    03-18-2008 - MM - Added ByVal References to the function parameters
	' </@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Function executeExcelQuery(ByVal sFilePath, ByVal sSQL) ' <@as> Recordset
	
	   Services.StartTransaction "executeExcelQuery" ' Timer Begin
	   Reporter.ReportEvent micDone, "executeExcelQuery Function", "Function begin"
	   
	   ' Variable Declaration / Initialization
	   Dim sConnection, oConnection, oResultSet
	   
	   'verifies that the passed parameters are not null or an empty string.
	   If IsNull(sFilePath) or sFilePath = "" or IsNull(sSQL) or sSQL = "" Then
	 			Reporter.ReportEvent micFail, "invalid parameter", "An invalid parameter was passed to the executeExcelQuery function."
	 			Set executeExcelQuery = Nothing
	 			Services.EndTransaction "executeExcelQuery" ' Timer End
	 			Exit Function
	   End If
	   
	   'creates a connection to the excel spreadsheet.
		 sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
									"Data Source=" & sFilePath & ";" & _
									"Extended Properties=""Excel 8.0;HDR=Yes; IMEX=1;""" 
		
		 'opens a connection.
		 Set oConnection = CreateObject("ADODB.Connection")
		 oConnection.Open sConnection
		
		 'executes the query if the database connection is open,
		 'otherwise the application closes and the test exits.
		 If oConnection.State = 1 Then
		 	Set oResultSet = CreateObject("ADODB.recordset")
		 	oResultSet.open sSQL, oConnection, 3,3
			Set executeExcelQuery = oResultSet
		 Else ' Connection not open
	   	Reporter.ReportEvent micFail, "connection state", "The database connection did not open."
	   	Set executeExcelQuery = Nothing
	 	 End If
	 	 
	   Reporter.ReportEvent micDone, "executeExcelQuery Function", "Function End"
	   Services.EndTransaction "executeExcelQuery" ' Timer End
	   
	End Function

	'<@comments>	
	'**********************************************************************************************
	' <@name>
	' executeCSVQuery
	' </@name>
	'
	' <@purpose>
	'	Creates a connection to the CSV file and get all of the data from the file as a 
	'           recordset object.
	' </@purpose>
	'
	' <@parameters>
	'	sFilePath = (ByVal) string - the path and filename of the CSV file.
	' </@parameters>
	'
	' <@return>
	'	An adodb recordset object containing the result of the csv file query.  
	'					 since the return item is an object, it must be set instead of assigned.
	'          Returns nothing if function failures
	' </@return>
	'
	' <@assumptions>
	'	environment variables have been initialized/created
	' </@assumptions>
	'
	' <@example_usage>
	'	Set oResultSet = db.executeCSVQuery("c:\file.csv")
	'
	'									If LCase(typeName(oResultSet)) <> "recordset" Then
	'										Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
	'
	'									ElseIf oResultSet.bof And oResultSet.eof Then
	'										Reporter.ReportEvent micFail, "invalid recordset", "The returned recordset contains no records."
	'
	'									Else
	'										Reporter.ReportEvent micPass, "valid recordset", "The returned recordset is valid and contains records."
	'										
	'										While Not oResultSet.eof
	'											sCode = oResultSet.fields(0).Value
	'											oResultSet.move(1)
	'										Wend
	'
	'									End If
	' </@example_usage>
	'
	' <@author>
	'	craig cardall
	' </@author>
	'
	' <@creation_date>
	'	07-30-2007
	' </@creation_date>
	'
	' <@mod_block>
	' 11-13-2007 - MM - Added Transaction Timers
	' 03-24-2008 - MM - Added logic to check the filename length is <= 64
	' 									Added Nothing return value if there is failure in the function
	'                   Added logic to verify the connection state of the 
	'                   ADODB.Connection and ADODB.RecordSet
	' </@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Function executeCSVQuery(ByVal sFilePath) ' <@as> Recordset
	
		Services.StartTransaction "executeCSVQuery" ' Timer Begin
		Reporter.ReportEvent micDone, "executeCSVQuery Function", "Function begin"
		
		Dim sConnection, oConnection, oResultSet, oFSO
		Dim sFileName, sFileType
		
		'verifies that the passed parameters are not null or an empty string.
		If IsNull(sFilePath) or sFilePath = "" Then
			Reporter.ReportEvent micFail, "invalid parameter", "An invalid parameter was passed to the executeCSVQuery function check passed parameters"
			Set executeCSVQuery = Nothing ' Return Value
			Services.EndTransaction "executeCSVQuery" ' Timer End
			Exit Function
		End If
		
		' Create a FileySystemObject
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		
		' Check to make sure file exists
		If oFSO.FileExists(sFilePath) Then
			' Get the File Type
			sFileType = oFSO.GetExtensionName(sFilePath)
			' Check file type name = csv
			If LCase(sFileType) = "csv" Then
				' Get the File Name
				sFileName = oFSO.GetFileName(sFilePath)
				
				' Parse sFilePath to separate the filename from the path
				sFilePath = Replace(sFilePath, sFileName, "", 1) ' Replace file name with nothing in sFilePath
				
				If Len(sFileName) <= 64 Then
					sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
												"Data Source=" & sFilePath & ";" & _
												"Extended Properties=""text;HDR=No; FMT=CSVDelimited"""
					sSQL = "select * from " & sFileName 'Set the SQL to get all data from the results spreadsheet
				Else ' FileName Parameter too long
					Reporter.ReportEvent micFail, "FileName Length", "The FileName '" & sFileName & "' is greater than 64 characters in length"
					Set executeCSVQuery = Nothing ' Return Value
					Services.EndTransaction "executeCSVQuery" ' Timer End
					Exit Function
				End If
				
				' opens a connection.
				Set oConnection = CreateObject("ADODB.Connection")
				oConnection.Open sConnection
				
				' Check the ADODB.Connection State, if open continue logic
				If oConnection.State = 1 Then
					
					' executes the query
					Set oResultSet = CreateObject("ADODB.recordset")
					oResultSet.open sSQL, oConnection, 3,3
					
					' Check the ADODB.Recordset connection state, if open continue logic
					If oResultSet.State = 1 Then
						Set executeCSVQuery = oResultSet ' Return Value
					Else ' Recordset connection closed
						Reporter.ReportEvent micFail, "ADODB Recordset", "Recordset Connection was not opened"
						Set executeCSVQuery = Nothing ' Return Value
					End If
				
				Else ' DB Connection Closed
					Reporter.ReportEvent micFail, "ADODB Connection", "DB Connection was not opened"
					Set executeCSVQuery = Nothing ' Return Value
				End If
				
			Else ' Incorrect file Type
				Reporter.ReportEvent micFail, "Invalid Parameter", "A CSV file is expected, not a '" & sFileType & "' when using the executeCSVQuery function check passed parameters"
				Set executeCSVQuery = Nothing ' Return Value
			End If
		Else ' File Does Not Exist
			Reporter.ReportEvent micFail, "Invalid Parameter", "File '" & sFilePath & "' does not exist when using the executeCSVQuery function check passed parameters"
			Set executeCSVQuery = Nothing ' Return Value
		End If
		
		' Clear Object Variables
		Set oFSO = Nothing
	
		Reporter.ReportEvent micDone, "executeCSVQuery Function", "Function End"
		Services.EndTransaction "executeCSVQuery" ' Timer End
	   
	End Function
	
	'<@comments>
	'**********************************************************************************************
	' <@name>
	'	executeMySQLQuery
	' </@name>
	'
	' <@purpose>
	'		To create a connection to the database and execute a sql statement
	' </@purpose>
	'
	' <@parameters>
	'        sQuery = (ByVal) String - A SQL statement to query a database
	'        sSchema = (ByVal) String - Name of the schema you want to execute against,
	'                                   this is case sensitive
	' </@parameters>
	'
	' <@return>
	'  				 An adodb recordset object containing the result 
	'          of the database query.  Since the return item is
	'          an object it must be set instead of assigned.
	'          Returns nothing if function failures
	' </@return>
	'
	' <@assumptions>
	'    Must have the MySQL ODBC Driver Installed on the machine
	'                      Can get the latest version from www.mysql.com
	'               Machine that uses this functions has rights to access the
	'               MySQL Database Server
	'               Environment Variables have been initialized/created
	'                 Environment("MYSQL_DB_SERVER")
	'                 Environment("MYSQL_DB_PORT")
	'                 Environment("MYSQL_SCHEMA_NAME")
	'                 Environment("MYSQL_SCHEMA_PSWD")
	' </@assumptions>
	'    
	' <@example_usage>
	'  set oRS = db.executeMySQLQuery("select * from table", "sys")	
	'
	'									If LCase(typeName(oRS)) <> "recordset" Then
	'										Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
	'
	'									ElseIf oRS.bof And oRS.eof Then
	'										Reporter.ReportEvent micFail, "invalid recordset", "The returned recordset contains no records."
	'
	'									Else
	'										Reporter.ReportEvent micPass, "valid recordset", "The returned recordset is valid and contains records."
	'										
	'										While Not oRS.eof
	'											sCode = oRS.fields(0).Value
	'											oRS.move(1)
	'										Wend
	'
	'									End If
	' </@example_usage>
	'
	' <@author>
	'		Mike Millgate
	' </@author>
	'
	' <@creation_date>
	'		04-07-2008
	' </@creation_date>
	'
	' <@mod_block>
	' </@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Function executeMySQLQuery(ByVal sQuery, ByVal sSchema) ' <@as> Recordset
	
		On Error Resume Next
		
		Services.StartTransaction "executeMySQLQuery" ' Timer Begin
	  Reporter.ReportEvent micDone, "executeMySQLQuery Function", "Function begin"
	   
	  ' Variable Declaration / Initialization
	  Dim oConnection, sConnectionString
	   
	  ' Check to verify passed parameters that they are not null or an empty string
	  If IsNull(sQuery) or sQuery = "" or IsNull(sSchema) or sSchema = "" Then
	  	Reporter.ReportEvent micFail, "Invalid Parameters", "Invalid parameters were passed to the executeMySQLQuery function check passed parameters"
	  	Set executeMySQLQuery = Nothing ' Return Value
	  	Services.EndTransaction "executeMySQLQuery" ' Timer End
	  	Exit Function
	  End If
	  
	  sConnectionString = "Driver={MySQL ODBC 3.51 Driver};" _
	                      & "Server=" & Environment.Value("MYSQL_DB_SERVER") & ";" _
	                      & "Port=" & Environment.Value("MYSQL_DB_PORT") & ";" _
	                      & "Database=" & sSchema & ";" _
	                      & "User=" & Environment.Value("MYSQL_SCHEMA_NAME") & ";" _
	                      & "Password=" & Environment.Value("MYSQL_SCHEMA_PSWD") & ";" _
	                      & "Option=3;"
   
	  ' Create Connection to database
	  Set oConnection = CreateObject("ADODB.Connection") ' Connection Object
	  
	  ' Set the connection to the connection string
	  oConnection.ConnectionString = sConnectionString
	  
	  ' Set the Connection Timeout Value
	  oConnection.ConnectionTimeout = 0
	    
	  ' Opens Database Connection
	  oConnection.Open
	  
	  If Err.Number <> 0 Then
	  	Reporter.ReportEvent micFail, "Connection Error", "Error Encountered: " & Err.Number & " - " & Err.Description
	  	Set executeMySQLQuery = Nothing ' Return Value
	  Else ' Continue logic
	  	' Execute sql statement, if database connection is open
	  	If oConnection.State = 1 Then
	  		Set executeMySQLQuery = oConnection.execute(sQuery) ' Return Value
	  	Else ' Connection not open
	  		Reporter.ReportEvent micFail, "connection state", "The database connection did not open."
	  		Set executeMySQLQuery = Nothing ' Return Value
	  	End If
	  End If
	  
	  Reporter.ReportEvent micDone, "executeMySQLQuery Function", "Function End"
	  Services.EndTransaction "executeMySQLQuery" ' Timer End
	   
	End Function

End Class

'**********************************************************************************************
'*                            Class Instantiation                                         
'**********************************************************************************************
dim db

set db = new DBFunctions