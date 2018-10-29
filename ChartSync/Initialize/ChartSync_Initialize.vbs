'**********************************************************************************************
' Initialize.vbs for ChartSync
' Functions - See individual Function Headers for
'             a detail description of function functionality

'       init_CloseExtraBrowsers
'       init_openBrowser
'       init_RegistryKeyExists
'        
'**********************************************************************************************

'*******************************************************************************************************
'  Purpose: Configures QTP Options/Settings
'			
'  Start: This is the first function/action to be called for any test.
'
'  End: QuickTest Pro is now configured to run the test correctly and
'       I.E. Browser Window is open to Application URL.
'
'  Assumptions: None
'
'  Parameters:  None-+
'
'  Notes:  None
'
'  Calls: init_CloseExtraBrowsers, init_openBrowser, init_RegistryKeyExists, init_SecurityAlert Functions or sub routines
'
'  Author: MM
'
'  Date: 12/14/2009 (Created)
'
'  Modifications: Hung Nguyen 12/12/10 Modified for ChartSync
'				  Sujatha Kandikonda 03/09/2011 Added Environment("BROWSER_TITLE") 
'				  Sujatha Kandikonda 03/09/2011 Added	Environment("TEST_BROWSER_URL")
'                 Hung Nguyen 7/18/11 - correct iStarrAppPath and added global environment var. tnsnames for db connection
'				  Govardhan Choletti 09/26/2011 - Added Environment paths for Tools(Putty, WinScp,Batch Job Templates)
'				  Govardhan Choletti 09/28/2011 - Added Environment Variables for Unix Server, ChartSync_ShellPath, Destination Folder
'				  Govardhan Choletti 05/09/2012 - Added Offshore User Roles, Configured for STG, UAT, SYS Enviroments
'                 Hung Nguyen 11/13/2012 - Added env. variable Environment("RESEARCH")="OnResearch" - the onshore research account 
'                 Hung Nguyen 11/15/12 - Added env. variable Environment("RESEARCH_CODER")="OnResCoder" - The onshore research coder account
'                 Hung Nguyen 1/14/13 Added 'Err.Clear' statement line 284 due to unknown error file loading.
'*******************************************************************************************************
Option Explicit
Services.StartTransaction "Initialize" ' Timer Begin   
Reporter.ReportEvent micDone, "Initialize", "Initialize Begin"
Dim iStarrAppPath,oCurTS,sTSNAME,sRunEnvironment,qtApp,bOpenBrowser,aResFileList,sResFile,sTestName
iStarrAppPath="\\nas00582pn\iSTARR\Automation\Application\ChartSync"	'framework file location

'************************************'Get test set name and set test environment (e.g., test/production/development)****************
'**************************************************************************************************************************************************
If QCUtil.IsConnected Then ' connected to QC
	If QCUtil.QCConnection.ProjectConnected Then ' connected to QC project

		'get test set name from QC  					
		Set oCurTS = QCUtil.CurrentTestSet ' set testSet Object - QuicTest must connected to Quality Center    			   
		If LCase(TypeName(oCurTS)) = "nothing" Then	'test obj not found - error connecs to QC?
				sTSNAME = "nothing"
		Else ' test obj found - QC connected successfully.
				sTSName = oCurTS.Name
		End If
		Set oCurTS = Nothing
   
		' set environment base on the test set name retrieved from QC (TEST, PROD, DEV)
		If InStr(1, LCase(sTSName), "_prod", 1) > 0 Then
			sRunEnvironment = "PROD"
		ElseIf InStr(1, LCase(sTSName), "_dev", 1) > 0 Then
			sRunEnvironment = "DEV"
		ElseIf InStr(1, LCase(sTSName), "_stg", 1) > 0 Then
			sRunEnvironment = "STG"			' Test the Scripts in STAGE Environment
		ElseIf InStr(1, LCase(sTSName), "_uat", 1) > 0 Then
			sRunEnvironment = "UAT"		' Test the Scripts in UAT Environment
		Else ' set default regardless
			sRunEnvironment = "TEST"		' Default running in SYS Environment
		End If
		Reporter.ReportEvent micInfo, "Quality Center Connection", "Domain = " & QCUtil.QCConnection.DomainName _
	  	                              & vbNewLine & ". Project = " & QCUtil.QCConnection.ProjectName _
	  	                              & vbNewLine & ". Testset = " &sTSNAME _
	  	                              & vbNewLine & ". Run Environment = " &sRunEnvironment	
	Else ' Not Connected to MQC Project
		Reporter.ReportEvent micDone, "Quality Center Connection", "Not Connected to MQC Project - Default to Test Environment"
		sRunEnvironment = "TEST" 	'set default regardless
	End If

Else ' Not Connected to MQC
	sRunEnvironment = "TEST"	' set default regardless
	Reporter.ReportEvent micInfo, "Quality Center Connection", "Not Connected to MQC. Default to Environment '" &sRunEnvironment &"'"
End If

' Get the Currently Running Test Name
sTestName = Environment.Value("TestName")
Reporter.ReportEvent micInfo, "Test name = " &sTestName,""


'************************************CREATE GLOBAL ENVIRONMENT HERE***************************************************************
'******************************************************************************************************************************************************
Environment("APP") = "CS" ' Acronym for the application
Environment("BROWSER_TITLE") = "Optum ChartSync"   ' Prefix to Browser Title
Environment("RUN_ENV") = sRunEnvironment ' Used to specify which environment the test is running against
'Environment("LOGFILE") = ""	'This will be used later to store a textstream object for a user log file. The textstream is created elsewhere.

' Test Environemnt URL's
If sRunEnvironment = "STG" Then
	'Environment("TEST_BROWSER_URL") = "http://ravasstage.ingenix.com/ravas/faces/unsecure/login.xhtml"  ' Testing STAGE URL
	Environment("TEST_BROWSER_URL") = "https://stagechartsync.optum.com/ravas/faces/unsecure/login.xhtml"  ' Testing STAGE URL
ElseIf sRunEnvironment = "UAT" Then
	Environment("TEST_BROWSER_URL") = "http://apspt0047:19080/ravas/faces/unsecure/login.xhtml"  ' Testing UAT URL
Else
	Environment("TEST_BROWSER_URL") = "http://apsp9108:29080/ravas/faces/unsecure/login.xhtml"  ' Testing SYS URL
End If
Reporter.ReportEvent micInfo, "Browser URL = " &Environment("TEST_BROWSER_URL"),""

'Environment("PROD_BROWSER_URL") = "http://www.production.url.com"' Production URL
'Environment("DEV_BROWSER_URL") = "http://www.development.url.com" ' Development URL
Environment("CHRT_REV_MGR") = "skcm"	'Chart review Manager
Environment("INTAKE1") = "skintake1"	'Intake Level 1 
Environment("INTAKE2")= "skintake2"     'Intake Level 2 
Environment("INTAKESUP")= "skintakesup"     'Intake Supervisor
Environment("CODER1") = "skcoder1"	  'Coder Level 1
Environment("CODER2") = "skcoder2"     'Coder Level 2
Environment("CODERSUP") = "skcodersup" ' Coder Supervisor
Environment("SYSADMIN")= "sksysadmin"     'System admin
Environment("CVANALYST")= "automation.cvanalyst"     'CV Analyst
Environment("CVMANAGER")= "skcvmgr"     'CV Manager
Environment("CVSUP")= "skcvsup"     'CV Supervisor
Environment("QAAUD")= "skqaaud"			'QA Auditor
Environment("QASUP")= "qasuptest1"						'QA Supervisor
Environment("QAMGR")= "qamgrtest1"						'QA Manager
Environment("CVQAMANAGER")= "skcvqamgr"     'CVQA Manager
Environment("CVQASUP")= "skcvqasup"     'CVQA Supervisor
Environment("CVQAAUDITOR")= "automation.cvqaaud"	'"mcvqaaud"     'CVQA Auditor
Environment("ACTMGR1")= "skactmgr1"     'Account Manager 1
Environment("ACTMGR2")= "skactmgr2"     'Account Manager 2
Environment("PRJANALYST")= "skPrjAnalyst"     'Project Analyst
Environment("PRJREQUESTOR")= "skprjreq"     'Project Requestor
Environment("HPCLIENT1")= "skhpclient1"     'HP Client Access 1
Environment("HPCLIENT2")= "skhpclient2"     'HP Client Access 2
Environment("ADMIN")= "admin"     					'Admin
Environment("RESEARCH")= "OnResearch"		'Onshore Research account
Environment("RESEARCH_CODER")= "OnResCoder"		'Onshore Research Coder account
Environment("RESEARCH_SUP")= "OnResSuper"		'Onshore Research supervisor account


' OFFSHORE USER ROLES
Environment("OFFINTAKE1") = "OffAutoIntake1"	'Offshore Intake Level 1 
Environment("OFFINTAKE2")= "OffAutoIntake2"     'Offshore Intake Level 2 
Environment("OFFINTAKESUP")= "OffAutoIntakeSup"     'Offshore Intake Supervisor
Environment("OFFCODER1") = "OffAutoCoder1"	  'Offshore Coder Level 1
Environment("OFFCODER2") = "OffAutoCoder2"     'Offshore Coder Level 2
Environment("OFFCODERSUP") = "OffAutoCoderSup" 'Offshore Coder Supervisor
Environment("OFFQAAUD")= "OffAutoQaAud"			'Offshore QA Auditor
Environment("OFFQASUP")= "OffAutoQaSup"		'Offshore QA Supervisor
Environment("OFFQAMGR")= "OffAutoQaMgr"		'Offshore QA Manager
Environment("OFFCVANALYST")= "OffAutoCvAnalyst"     'Offshore CV Analyst
Environment("OFFCVSUP")= "OffAutoCvSup"     'Offshore CV Supervisor
Environment("OFFCVMANAGER")= "OffAutoCvMgr"     'Offshore CV Manager
Environment("OFFCVQAAUDITOR")= "OffAutoCvqaAud"		'Offshore CVQA Auditor
Environment("OFFCVQASUP")= "OffAutoCvqaSup"     'Offshore CVQA Supervisor
Environment("OFFCVQAMANAGER")= "OffAutoCvqaMgr"     'Offshore CVQA Manager
' COMMON PASSWORD for All User Roles
Environment("CHARTSYNC_PWD")= "password"		'Password
Environment("RESET_PWD")= "Password-125"		'Password Reset

Environment("CONNECTION_TYPE") = "ADODB.Connection" '  Used for ADO and/or ODBC  Database connections
Environment("CONNECTION_PROVIDER") = "OraOLEDB.Oracle" '  Used for the ADO and/or ODBC Database Connections
' SET DB for different Environments
If sRunEnvironment = "STG" Then
	Environment("XTSN") = "RAVSTG"		'tnsname for STAGE Environment
ElseIf sRunEnvironment = "UAT" Then
	Environment("XTSN") = "RAVUAT"		'tnsname for UAT Environment
Else
	Environment("XTSN") = "RAVSYS"		'tnsname for SYSTEM Environemnt Oracle 11G
	'Environment("XTSN") = "RAVTEST"		'tnsname for SYSTEM Environemnt Oracle 10G
End If
Environment("SCHEMA_NAME") = "UHG_000746373"	'"UHG_000720722"	'"UHG_000746373"	'"RAVAS" '  Used for the ADO and/or ODBC Database Connections - Schema User Name
Environment("SCHEMA_PSWD") = "Gov~1313"		'"Jehovahnissi2*"	'"Gov~1234"		'"RAVASTEST" '  Used for the ADO and/or ODBC Database Connections - Schema Password
Environment("DB_INCREMENTOR") = 1 ' Used to step through data retrieved from the database
Environment("SYS_SCHEMA_NAME") = "UHG_000746373"	'"UHG_000720722"	'"UHG_000746373"	'"RAVAS"	'"RAVAS2" '  Used for the ADO and/or ODBC Database Connections - Schema User Name
Environment("SYS_SCHEMA_PSWD") = "Gov~1313"		'"Jehovahnissi2*"	'"Gov~1234"		'"RAVASTEST"	'"ravas2" '  Used for the ADO and/or ODBC Database Connections - Schema Password

'**************************** Batch Job Process *************************************
'Environment("winscppath")  = "\\Nas00582pn\istarr\Automation\Application\IRADS\exe\WinSCP\"
'Environment("puttyexepath")  = "\\Nas00582pn\istarr\Automation\Application\IRADS\exe\Putty\putty.exe"
'Environment("CurDate") = Replace(Date, "/", "-")
'If Environment.Value("RunEnv") = "TEST" Then
'	Environment("Rule_Template_Path") = "\\Nas00582pn\istarr\Automation\Application\IRADS\RuleTemplates\SYS\"&Environment.Value("TestName")&"\"
'	Environment("Rules_Path") = "\\Nas00582pn\istarr\Automation\Application\IRADS\Rules\SYS\"
'ElseIf Environment.Value("RunEnv") = "TEST1" Then
'	Environment("Rule_Template_Path") = "\\Nas00582pn\istarr\Automation\Application\IRADS\RuleTemplates\SYS2\"&Environment.Value("TestName")&"\"
'	Environment("Rules_Path") = "\\Nas00582pn\istarr\Automation\Application\IRADS\Rules\SYS2\"
'End If

Environment("winscppath")  = "\\Nas00582pn\istarr\Automation\Application\ChartSync\ExeFiles\WinSCP\"
Environment("puttyexepath")  = "\\Nas00582pn\istarr\Automation\Application\ChartSync\ExeFiles\Putty\putty.exe"
Environment("CurDate") = Replace(Date, "/", "-")
'If Environment.Value("RUN_ENV") = "TEST" Then
If Environment.Value("RUN_ENV") = "STG" Then
	Environment("Batch_Job_Path") = "\\Nas00582pn\istarr\Automation\Application\ChartSync\BatchJobTemplates\STG\"&Environment.Value("TestName")&"\"
	Environment("Jobs_Path") = "\\Nas00582pn\istarr\Automation\Application\ChartSync\BatchJobs\STG\"
Else
	Environment("Batch_Job_Path") = "\\Nas00582pn\istarr\Automation\Application\ChartSync\BatchJobTemplates\QA\"&Environment.Value("TestName")&"\"
	Environment("Jobs_Path") = "\\Nas00582pn\istarr\Automation\Application\ChartSync\BatchJobs\QA\"
End If

' ************** UNIX SERVER DETAILS ******************
'Environment("UnixServer") = "apsp9120"
'Environment("PortNumber") = "22"
'Environment("UnixUserName") = "sbolem"
'Environment("UnixPassword") = "juhisa"
'SET UNIX BOX Server Details
If sRunEnvironment = "STG" Then
	Environment("UnixServer") = "apsp9307" 		' Unix Box for Stage Environment
ElseIf sRunEnvironment = "UAT" Then
	Environment("UnixServer") = "apsp8472" 		' Unix Box for UAT Environment
Else
	Environment("UnixServer") = "apsp9052" 		' Unix Box for SYSTEM Environment
End If
Environment("PortNumber") = "22"
Environment("UnixUserName") = "ostdvu"
Environment("UnixPassword") = "lad758as"


If Environment.Value("RUN_ENV") = "STG" OR Environment.Value("RUN_ENV") = "TEST" Then 'Executing in Test Environment
'	Environment("DestFolder") = "/INTERFACES_IRADS/irads/data/proc/in"
'	Environment("IRADS_shellexecpath") = "/INTERFACES_IRADS/irads/usr/batch/bin"
	Environment("DestFolder") = "/ravas/data/proc/in"
	Environment("ChartSync_shellexecpath") = "/ravas/usr/batch/bin"
End If

' *********************************Init QuickTest **********************************************************************************************************
'******************************************************************************************************************************************************
Set qtApp = CreateObject("QuickTest.Application")
qtApp.Visible = True 	'Sets QTP to maximized window mode

'Close all open browsers except Mercury Quality Center and Mercury Support
bOpenBrowser = False
init_CloseExtraBrowsers ' Calls Local Function

''Configuring QTP in General Options (Tools >> Options)
qtApp.Options.Run.RunMode = "Fast"    		   'Set QTP run mode: Fast or Normal
qtApp.Options.Run.ViewResults = "False"		'Set QTP to view results after run (e.g., True=view, False=no view)
qtApp.Options.Run.CaptureForTestResults = "OnWarning"	'capture Active Screen infor in test results on warning and fail

'Set QTP where to look for any tests, actions, or files. Allows files to be loaded without having to include their full paths
qtApp.Folders.RemoveAll
qtApp.Folders.Add(iStarrAppPath)

''Configuring QTP in Test Settings (Test >> Settings)
qtApp.Test.Settings.Run.DisableSmartIdentification = "True" 						'Set QTP whether or not to use Smart Identification
qtApp.Test.Settings.Run.OnError = "NextStep"													'Set  QTP what to do when it encounters a script error
qtApp.Test.Settings.Launchers("Windows Applications").Active = False	'Configure the Windows application to run on any open application

''Configuring Record and Run Settings (Automation >> Record and Run Settings...)
qtApp.Test.Settings.Launchers("Web").Active = False

' 'Opens Application URL
''Launches IE and navigates to Application URL
Select Case LCase(sRunEnvironment)
'	Case "prod"
'		init_openIEBrowser(Environment.Value("PROD_BROWSER_URL")) ' Calls Sub Function passing in the URL set in the environment variable
	Case "test", "stg", "uat"
		If init_openIEBrowser(Environment.Value("TEST_BROWSER_URL")) Then ' Calls Sub Function passing in the URL set in the environment variable
			Environment.Value("DB") = "sys"
			Reporter.ReportEvent micPass, "init_openIEBrowser", "The init_openIEBrowser Function call was successful"
		Else ' Not Successful
			Reporter.ReportEvent micFail, "init_openIEBrowser", "The init_openIEBrowser Function call was not successful, Exiting Test"
			Services.EndTransaction "Initialize" ' Timer End
			ExitTest ' Exits the test
		End If
'	Case "dev"
'		init_openIEBrowser(Environment.Value("DEV_BROWSER_URL")) ' Calls Sub Function passing in the URL set in the environment variable
	Case Else
		Reporter.ReportEvent micFail, "Test Environment", _
			"No valid Test Environment was selected. Edit the script with a valid environment "
		Services.EndTransaction "Initialize" ' Timer End
		ExitTest ' Exits the test
End Select

' **********************************************
' Load all external function files if the testname does not contain "_debug"
' If it does contain "_debug" then function files
' Need to be loaded via the test script with either ExecuteFile or as Resource Files
' ExecuteFile function modified to LoadFunctionLibrary
' **********************************************
If Not InStr(1, LCase(sTestName), "_debug", 1) > 0 Then
	Reporter.ReportEvent micDone, "Function File Load", "Test not in Debug mode - Function Files loaded via Initialize Script"
	
	aResFileList = Array("login.vbs","logout.vbs","ErrorHandler.vbs","ajaxSync.vbs","objFunctions.vbs","utilFunctions.vbs", "dbFunctions.vbs","ChartSyncBatchJobs.vbs")
	Err.Clear
	For each sResFile in aResFileList
		' TILL QTP 10, we have used the below specified method to execute
		'ExecuteFile(sResFile)		' Attempt to load the file into memory
		
		'QTP ALM 11 New Method introduced LoadFunctionLibrary 
		LoadFunctionLibrary iStarrAppPath &"\"&sResFile
		
		' Check to see if an Error was encountered while trying to execute the file
		If Err.Number <> 0 Then
			Reporter.ReportEvent micFail, "LoadFunctionLibrary", "LoadFunctionLibrary logic for file '" & sResFile & "' was not successful" _
			                                             & vbNewLine & "Error Encountered: " & Err.Number & vbNewLine & Err.Description _
			                                             & vbNewLine & "Exiting Test"
			Services.EndTransaction "Initialize" ' Timer End
			ExitTest ' Exits the test
		Else ' LoadFunctionLibrary Successful
			Reporter.ReportEvent micPass, "LoadFunctionLibrary", "LoadFunctionLibrary logic for file '" & sResFile & "', was successful"
		End If
	Next
Else
	Reporter.ReportEvent micDone, "Function File Load", "Test in Debug mode - Function Files not loaded via Initialize Script"
End If

Set qtApp = Nothing	'destroy obj

Reporter.ReportEvent micDone, "Initialize", "Initialize End"
Services.EndTransaction "Initialize" ' Timer End


' ***********************
' LOCAL FUNCTIONS
' ***********************
Sub init_CloseExtraBrowsers()
   '**********************************************************************************************
   ' Purpose:  This subroutine closes all open browsers except Mercury Quality Center and Mercury Support.
	 '           Other browsers can be added to the list by creating more ElseIf clauses
	 '           in the reverse order For..Next loop below
   ' Parameters: None
   ' Returns: Nothing
   ' Assumptions:  None
   ' Example Usage: init_CloseExtraBrowsers()
   ' Calls:  init_RegistryKeyExists function
   ' Author: Mike Millgate
   ' Date: 03/02/2006
   ' Modifications:  
   ' 
   '**********************************************************************************************
   Services.StartTransaction "init_CloseExtraBrowsers" ' Timer Begin
   Reporter.ReportEvent micDone, "init_CloseExtraBrowsers Sub", "Sub Begin"

	Dim varBrowsers, iCounter, varHwnd, Flag
	Dim varLastHwnd, varOpenBrowser, iOldTimeout
	iCounter = 0 
	Flag = 1
	 
	 ' Gets the current Registry Key value for Microsoft Internet Explorer Window Title
	 Dim oWshShell, sTemp, counter
	 
	 Const regKey = "HKCU\Software\Microsoft\Internet Explorer\Main\Window Title"
	 Set oWshShell = CreateObject("Wscript.Shell") ' Create Shell object to read/write to Registry
	 
	 ' Check to see if registry Key exists, if it does get/store the value of the key
	 If init_RegistryKeyExists(regKey) Then
	 	sTemp  = Trim(oWshshell.RegRead(regKey)) ' Get/Read current key value
	 End If 
	 Set oWshShell = nothing ' Releasing object
	 
	 iOldTimeout = qtApp.Test.Settings.Run.ObjectSyncTimeOut	'get current time out
	 qtApp.Test.Settings.Run.ObjectSyncTimeOut = 1000					'set time out to 1 sec waiting to find obj
	 
	 Set varBrowsers = CreateObject("Scripting.Dictionary")
	 While (Window("regexpwndclass:=IEFrame","index:=" & iCounter).Exist And Flag)
	 	'wait 1 	'this short wait may be needed but for now it seems to work fine without it
	 	varHwnd = Window("regexpwndclass:=IEFrame","index:=" & iCounter).getroproperty("Hwnd") 
		If (varLastHwnd = varHwnd) Then
			Flag = 0
		Else
			varBrowsers.Add CStr(varBrowsers.Count), varHwnd
			iCounter = iCounter+1
		End If
		varLastHwnd = varHwnd
	 Wend
	 
	 'Loop to close all opened objects
	 For iCounter = varBrowsers.Count-1 to 0 step -1 'close the opened objects in reverse order
	 	varHwnd = varBrowsers.Item(CStr(iCounter))
	 	Wait 1	'gimme 1 sec
	 	varOpenBrowser = Trim(Window("regexpwndclass:=IEFrame","index:=" & iCounter).getroproperty("title"))
	 	If Instr(1,varOpenBrowser,"HP Quality Center",1) Then
	 		Reporter.ReportEvent micDone, "HP Quality Center", "Keep the HP Quality Center browser window open."
	 		bOpenBrowser = True	'Indicates that a browser has been left open so no additional browsers will be opened by this script	 		
		ElseIf Instr(1,varOpenBrowser,"HP Application Lifecycle Management",1) Then
	 		Reporter.ReportEvent micDone, "HP Application Lifecycle Management", "Keep the HP Application Lifecycle Management browser window open."
	 		bOpenBrowser = True	'Indicates that a browser has been left open so no additional browsers will be opened by this script	 		
	 	ElseIf Instr(1,varOpenBrowser,"Optum ChartSync",1) Then	
	 		'propertly logout prior to close the Browser obj
	 		If Browser("name:=Optum ChartSync").Link("html id:=headerFormNav:logoffLink").Exist(1) Then
	 			Browser("name:=Optum ChartSync").Link("html id:=headerFormNav:logoffLink").Click
	 			
	 			If Browser("name:=Optum ChartSync").WebButton("html id:=confirmLogoutModalPanelForm:confirmLogoutModalPanelConfirmButton").Exist(3) Then 
	 				Browser("name:=Optum ChartSync").WebButton("html id:=confirmLogoutModalPanelForm:confirmLogoutModalPanelConfirmButton").Click
	 			End If
	 		End If
	 		Reporter.ReportEvent micDone, varOpenBrowser & " Closed", "Browser closed: " & varOpenBrowser
	 		Window("hwnd:=" & varHwnd).Close	 		
	 	Else
	 		Reporter.ReportEvent micDone, varOpenBrowser & " Closed", "Browser closed: " & varOpenBrowser
	 		Window("hwnd:=" & varHwnd).Close
	 	End If
	 Next
	 
	 qtApp.Test.Settings.Run.ObjectSyncTimeOut = iOldTimeout	'restore orig. time out value
	 
	 Reporter.ReportEvent micDone, "init_CloseExtraBrowsers Sub", "Sub End"
	 Services.EndTransaction "init_CloseExtraBrowsers" ' Timer End
End Sub

Function init_openIEBrowser(ByVal sUrl) 
	'************************************************************************************************************************
	'Purpose: IE - Navigate to the URL specified
	'Note: the global env. variable 'Environment("BROWSER_OBJ")' is created if function call was successful
	'Parameters: sUrl = URL address
	'Returns: True/False
	'Calls: None
	'Usage:  Call  init_openIEBrowser("http://apsp9108:29080/ravas/faces/unsecure/login.xhtml")  
	'Created by: Hung Nguyen -10/26/2012
	'Modified: Hung Nguyen 1/9/13 - Use obj.Exist w/30secs max. timeout.
	'Modified: Govardhan Choletti 6/5/13 - Modify Systemutil.Run to COM Object because QTP with ALM 11 opens two Browsers 
	'************************************************************************************************************************
	Services.StartTransaction "init_openIEBrowser"
	init_openIEBrowser = False ' init Return Value

	' Check if parameter is not an empty string
	If sUrl = "" Then
		Reporter.ReportEvent micFail, "Invalid Parameters", "Parameter can't be empty."
		Services.EndTransaction "init_openIEBrowser" 
		Exit Function
	End If

	Dim oBrowserApp,oIExplore,sName,cnt
	
	' Launch Internet Explorer and navigate to the URL specified
	Err.Clear
	On Error Resume Next 
	'SystemUtil.Run("IEXPLORE.EXE"),sUrl,,,3
	'USE COM OBJECT To Open Browser as SystemUtil.Run is opening 2 Browsers
	Set oIExplore = CreateObject("InternetExplorer.Application")
	oIExplore.Visible = True
	oIExplore.Navigate sUrl
	While oIExplore.Busy
	Wend
	
	If Not Browser("creationtime:=0").Exist(30) Then	'30secs max. 
		reporter.reportevent micFail,"init_openIEBrowser","Launching IE Browser failed. Time out issue?"
		Exit Function
	End If 
	
	Browser("creationtime:=0").WaitProperty "name",micNotEqual(""),60000
	sName=Trim(Browser("creationtime:=0").GetROProperty("name"))
 
	'if error 	
	If InStr(1,sName,"HTTP 404 Not Found",1) > 0 Then
		reporter.reportevent micFail,"Navigation error - HPPP 404 Not Found","" 
		Exit Function
	ElseIf StrComp(sName,Environment("BROWSER_TITLE"),1) = 0 Then
		Set oBrowserApp = Browser("title:="&sName)
		Environment("BROWSER_OBJ") = oBrowserApp ' Set global Environment Variable for the Browser obj
		init_openIEBrowser = True ' Return Value
		
		Reporter.ReportEvent micPass, "Launching Browser Window", "Navigate to the URL '" &sUrl &"' was successful." _
		                     &vbNewLine &"Browser name '" &sName &"' found."
	Else
		Reporter.ReportEvent micFail, "Launching Browser Window", "Unable to retrieve the Browser name after navigation. Value retrieved '" &sName &"'"
	End If
	
	On Error Goto 0 'reset
	Set oBrowserApp = Nothing
	Set oIExplore = Nothing
	Services.EndTransaction "init_openIEBrowser"
End Function

Function init_RegistryKeyExists(RegistryKey)
   '**********************************************************************************************
   ' Purpose:  Checks to see if a registry key exists
   ' Parameters:
   '        RegistryKey = String - Name of the key and key path to be verified
   ' Returns: Boolean Value (True/False)
   '            True -  If found
   '            False - If not found or other function errors
   ' Assumptions:  None
   ' Example Usage: init_RegistryKeyExists("HKCU\Software\Microsoft\Internet Explorer\Main\Window Title")
   ' Author: Mike Millgate
   ' Date: 09-26-2007
   ' Modifications:  
   '**********************************************************************************************
   Services.StartTransaction "init_RegistryKeyExists" ' Timer Begin
   Reporter.ReportEvent micDone, "init_RegistryKeyExists", "Function Begin"
   
   ' Check to verify passed parameters that they are not null or an empty string
   If IsNull(RegistryKey) or RegistryKey = "" Then
   		Reporter.ReportEvent micFail, "Invalid Parameters", "Invalid parameters were passed to init_RegistryKeyExists function check passed parameters"
   		init_RegistryKeyExists = False ' Return Value
   		Services.EndTransaction "init_RegistryKeyExists" ' Timer End
   		Exit Function
   End If

   On Error Resume Next
   WshShell.RegRead RegistryKey   'Try reading the key
   
   'Catch the error
   Select Case Err
	Case 0 ' Error Code 0 = 'success'
   		init_RegistryKeyExists = True ' Return Value
   	Case  1 'This checks for the (Default) value existing (but being blank); as well as key's not existing at all (same error code)
      ErrDescription = Replace(Err.description, RegistryKey, "")   		'Read the error description, removing the registry key from that description
      Err.clear      'Clear the error
      
      'Read in a registry entry we know doesn't exist (to create an error description for something that doesnt exist)
      WshShell.RegRead "HKEY_ERROR\"
       
      'The registry key exists if the error description from the HKEY_ERROR RegRead attempt doesn't match the error
      'description from our RegistryKey RegRead attempt
      If (ErrDescription <> Replace(Err.description, "HKEY_ERROR\", "")) Then
      	init_RegistryKeyExists = True ' Return Value
      Else
      	init_RegistryKeyExists = False ' Return Value
      End If
   	Case Else 'Any other error code is a failure code
   	  init_RegistryKeyExists = False ' Return Value
   End Select
   
   'Turn error reporting back on
   On Error Goto 0
   
   Reporter.ReportEvent micDone, "init_RegistryKeyExists", "Function End"
   Services.EndTransaction "init_RegistryKeyExists" ' Timer End
End Function