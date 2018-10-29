'===========================================================================================
' Purpose					 		: Call to DriverScript Function in CRAFT framework
' Author		   			 		   : Cognizant Technology Solutions
' Last Modification	  		 :  17-Feb-2010
'Last ModificationBY     : Teekam Singh Karki, Ram (229673) - Added Code to handle KEDS (6/1/2010), GH (2/8/10)
'===========================================================================================
'Variable Declaration
'===========================================================================================
	Dim fso, flInputTextFile
	Dim strCompName, blnExecute, blnReplaceHost
	Dim clsDriverScript, clsEnvironmentVariables
	Dim strGroupName,intRowCountMaster
	'Desc: To handle QC integration
	Dim strQCIntegrationArray()
	Dim ResultFileName()   
	Set fso = CreateObject("Scripting.FileSystemObject")

	' Desc: ScriptPath,  Store the complete path of Run Manager
	ScriptPath = Environment("TestDir")   

	' Desc: To handle QC integration
	DriveName = fso.GetDriveName(ScriptPath)
	DriveName = DriveName & "\"
	strControlPath=DriveName & "PaylessEComm_RegressionAutomation\CRAFT\"

	'' Desc: This below code will not execute when Folder structure exist in local system.
	'' Mean it will  download the script from QC , when folder CRAFT  framework is not setup in local drive.
	' Download framework on local machine if don't exist to run the scripts
	If not fso.FolderExists(strControlPath)  Then
		Call DownloadFrameworkOnMachine(strControlPath)
	End If


''Adding Code for  marking "N" and "Y" with Group master and test case to execute , 
'which pass through QCtest case at rtun time 
''====Start 

	If UCase ( Environment.Value ( "TestName") ) <> "RUNMANAGER" Then
		CurrentTestCaseName = QCUtil.CurrentTest.Name
		''CurrentTestCaseName =  Environment.Value ( "TestName")  ''testing both option

			Environment.Value("strCurrentTestName")=CurrentTestCaseName
	
			If UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "ORIG_QTP"  Then
				Environment.Value ( "ModuleName") =   "Original"
				strModuleName = "Original"
			ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "PERF_QTP"  Then
				Environment.Value ( "ModuleName") =   "Performance"
				strModuleName = "Performance"
			ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "SPER_QTP"  Then
				Environment.Value ( "ModuleName") =   "Sperry"
				strModuleName = "Sperry"

				'For KEDS - Ram 28-5-2010
			ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "KEDS_QTP"  Then
				Environment.Value ( "ModuleName") =   "Keds"
				strModuleName = "Keds"                
             'For Grasshoppers - 2-Aug-10
			ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "GRAS_QTP"  Then
				Environment.Value ( "ModuleName") =   "Grasshoppers"
				strModuleName = "Grasshoppers"			
			End If 

'' Why we are importing sheet from QC instead of local copy
		'DataTable.ImportSheet "[QualityCenter] Subject\PaylessEComm_RegressionAutomation\CRAFT\Business_Scripts\TestCase_Scripts\PAYLESS_ECOMM\Test_Data\Manual_Automation_Mapping.xls" , 1 ,1 
		DataTable.ImportSheet "C:\PaylessEComm_RegressionAutomation\CRAFT\Business_Scripts\TestCase_Scripts\PAYLESS_ECOMM\Test_Data\Manual_Automation_Mapping.xls" , strModuleName , 1 
		rowcount = DataTable.GetSheet("Global").GetRowCount 
		For i=1 To rowcount
			DataTable.SetCurrentRow(i)
			QC_TestCase_Name=DataTable.Value("QC_TestCase_Name","Global")
			If UCase( QC_TestCase_Name ) = UCase( CurrentTestCaseName ) Then 
				Automation_TestCase_Name = DataTable.Value("Automation_TestCase_Name","Global")
				Group_Name_Automation = DataTable.Value("Group_Name","Global")
				Exit For				   
			End If   
		Next		
		Group_Test_Case_Modify  "PAYLESS_ECOMM_GroupMaster" , Group_Name_Automation,"GroupName"
		Group_Test_Case_Modify  Group_Name_Automation , Automation_TestCase_Name ,"TestCase_Name"
	Else
		''Temp Value need to commnet  below line at run time execution  
		''When Execution in local system through Runmanager	for testing 
    	CurrentTestCaseName = "TestingScript_KEDS_QTP"  ''Temp for testing or local execution

		Environment.Value("strCurrentTestName")=CurrentTestCaseName				
		If UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "ORIG_QTP"  Then
			Environment.Value ( "ModuleName") =   "Original"
			strModuleName = "Original"
        ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "PERF_QTP"  Then
			Environment.Value ( "ModuleName") =   "Performance"
			strModuleName = "Performance"
        ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "SPER_QTP"  Then
			Environment.Value ( "ModuleName") =   "Sperry"
			strModuleName = "Sperry"
		ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "KEDS_QTP"  Then
			Environment.Value ( "ModuleName") =   "Keds"
			strModuleName = "Keds"
		ElseIf UCase ( Right (  CurrentTestCaseName  , 8 ) ) = "GRAS_QTP"  Then
				Environment.Value ( "ModuleName") =   "Grasshoppers"
				strModuleName = "Grasshoppers"		
		End If 

	End If 
''===== End 

 
	strProjectName = "PAYLESS_ECOMM"
	strAppType = "Web"
		  
	Reporter.Filter = rfDisableAll

	'Desc: To handle QC integration
	'Author: Cognizant Technology Solutions
	i=0
	'Checks for the Initial configuration file and executes all the class files.	
	If fso.FileExists(strControlPath & "Environment_SetUp\"&strProjectName&"\Setup.ini") Then
	
		Set flInputTextFile = fso.OpenTextFile(strControlPath &  "Environment_SetUp\"&strProjectName&"\Setup.ini",1)		
		strSetupData = Split(flInputTextFile.ReadAll,vbcrlf)
        flInputTextFile.Close

        ''Desc : For retriving the data from setup.ini file and execute those VBS file	
		For iCase = LBound(strSetupData) To UBound(strSetupData)
			
			strCase = Split(strSetupData(iCase),"=")
  			if UCase(Right(Trim(strCase(0)),4)) = "FILE" Then
				strCase(0) = "FILE"			
				
			'Desc: To handle QC integration defined in setup.ini file
			ElseIf UCase(Right(Trim(strCase(0)),2)) = "QC" Then
					strCase(0) = "QC"

        	ElseIf UCase(Right(Trim(strCase(0)),3)) = "DSN" Then  '' Not in used , Teekam 22 -Feb
				strCase(0) = "DSN"

			End If
				
			Select Case strCase(0)			
				Case "FILE"
                        If fso.FileExists(strControlPath & strCase(1)) Then
							strCase(1)=Cstr(strCase(1))
							ExecuteFile strControlPath & strCase(1)
                        End If  
							
				'Desc: To handle QC integration
				Case "QC"
							ReDim Preserve strQCIntegrationArray(i)
							strQCIntegrationArray(i)=CStr(strCase(1))
							 i=i+1

				Case "DSN"  '' Not in use Teekam 22-Feb 
							temp_hold=CStr(strCase(1))

			End Select
			
		Next
	Else
	    ''Desc: Enter the detail  of  error in QC 
		Reporter.Filter = rfEnableAll
		Reporter.ReportEvent 3,"Initial Input","Wrong file name entered."
		Reporter.Filter = rfDisableAll
		ExitRun
	End If


    'Instantiates all the objects.required for class defined in VBS file 
	Set clsReport = New Report
	Set clsUtilityScript= New UtilityScript
	Set clsEnvironmentVariables = New EnvironmentVariables
	Set clsInitScript = New InitScript	
	Set clsRepLoadScript= New RepositoryLoadScript
	Set clsRecoveryLoadScript= New RecoveryLoadScript
	Set clsDatabase_Module=New Database_Module

	'Desc: To handle QC integration
	Set clsQCIntegration_Module=New QCIntegration_Module

    strCntrlPath=strControlPath
	clsEnvironmentVariables.CntrlPath = strCntrlPath
	
    'Initializes all the environment parameters.
    clsEnvironmentVariables.Connection_String=temp_hold
	clsEnvironmentVariables.ProjectName = strProjectName
	clsEnvironmentVariables.AppType = strAppType
    clsEnvironmentVariables.MasterXLSPath = strControlPath & "Business_Scripts\TestCase_List\"&strProjectName&"\"& strProjectName & "_GroupMaster.xls"
	clsEnvironmentVariables.ControlPath = strControlPath + "Business_Scripts\TestCase_List\"&strProjectName
	clsEnvironmentVariables.TempFilePath = Environment("SystemTempDir")
	clsEnvironmentVariables.TempSummaryFilePath = Environment("ResultDir")
    clsEnvironmentVariables.RunTimeReportPath = strControlPath + "Framework_Scripts\Reports\"&strProjectName&"\Runtime_Reports"
	clsEnvironmentVariables.TestResultPath = strControlPath + "Framework_Scripts\Reports\"&strProjectName&"\Test_Results_Log"
	clsEnvironmentVariables.SummaryResultPath =  strControlPath + "Framework_Scripts\Reports\"&strProjectName&"\Summary"
	clsEnvironmentVariables.RunBy = Environment("UserName")   ''  Not in Use Teekam 22-Feb

	'Desc: To handle QC integration
	clsEnvironmentVariables.ServerNameQC = strQCIntegrationArray(0)
	clsEnvironmentVariables.UserNameQC = strQCIntegrationArray(1)
	clsEnvironmentVariables.PasswordQC = strQCIntegrationArray(2)
	clsEnvironmentVariables.DomainQC = strQCIntegrationArray(3)
	clsEnvironmentVariables.ProjectQC = strQCIntegrationArray(4)
	clsEnvironmentVariables.FolderNameQC = strQCIntegrationArray(5)
	clsEnvironmentVariables.TestSetPathQC = strQCIntegrationArray(6)
	clsEnvironmentVariables.TestSetNameQC = strQCIntegrationArray(7)
	clsEnvironmentVariables.QCUpdation = strQCIntegrationArray(8)
    blnQCUpdation = clsEnvironmentVariables.QCUpdation
	blnMultiUserExecution= CBool(strQCIntegrationArray(9))
	
	clsEnvironmentVariables.OverNightRun = strQCIntegrationArray(10)  '' Teekam 20 Jan 



   'Download Master excel sheet from QC if Multi user support is enabled
	If UCase ( Environment.Value ( "TestName") ) = "RUNMANAGER" Then  ''Teekam added for performance 28 Oct 
			If blnMultiUserExecution = True Then
					Call DownloadMasterExcel()
			End If
	End If


	'If the excel file exist then delete it. 'Create ECOMMTestResult excel file . This file is used to write test results
	''Teekam 28 Oct 
	If UCase ( Environment.Value ( "TestName") ) = "RUNMANAGER" Then  ''Teekam added for performance 28 Oct 
		ExcelFilePath= strControlPath & "Framework_Scripts\Reports\"& strProjectName &"\Test_Results_Log\"    
		ExcelFileWithPath = ExcelFilePath &  "ECOMMTestResults.xls"
		If (fso.FileExists(ExcelFileWithPath)) Then
				fso.deletefile(ExcelFileWithPath)
		End If
	End If 
	Set fso = Nothing


    ''Desc: This step is not  executable for local drive execution
	If UCase ( Environment.Value ( "TestName") ) = "RUNMANAGER" Then  ''Teekam added for performance improvement  29 Oct 
			If  blnQCUpdation Then
					clsEnvironmentVariables.TestResultExcelFile=ExcelFileWithPath
					Set clsEnvironmentVariables.UseExcelObject =  clsQCIntegration_Module.CreateExcelFile()
			End If
	End If 


    'SystemUtil.CloseDescendentProcesses 
    DataTable.AddSheet "Environment"	
	IntSheetNo=DataTable.GetSheetCount
	DataTable.ImportSheet strControlPath & "Environment_SetUp\"&strProjectName&"\" & strProjectName & "_Environment.xls" , "Sheet1" , "Environment"

''''-----------Teekam-----------------
'' Here we have to get the application URL from environment sheet and set into environment  variable.
'''-----------Teekam-----------------

	rowcount = DataTable.GetSheet("Environment").GetRowCount 
		For i=1 To rowcount
			DataTable.SetCurrentRow ( i )
			strEnvName =DataTable.Value ( "EnvironmentType" , "Environment" )	
			If UCase( strModuleName ) = UCase( strEnvName ) Then 
				AppURL = DataTable.Value ( "URL" , "Environment" )

				strIEBrowser  = DataTable.Value ( "IE" , "Environment" )
				strFirefoxBowser  = DataTable.Value ( "Firefox" , "Environment" )
				If UCase( Trim( strIEBrowser ) ) = "Y" Then
					Environment.Value ("BrowserName") = "IE"
				ElseIf UCase( Trim( strFirefoxBowser ) ) = "Y" Then
					Environment.Value ("BrowserName") = "Firefox"
				End If

				Exit For				   
			End If   	
		Next


''Added by Teekam 28 Oct for environment value 	
''Commneted by Teekam - 22-Feb 
	''strEnv=Ucase( Replace(DataTable("EnvironmentType","Environment")," ","")  )

    'Imports the Groupmaster excel to the global sheet of the run-time excel.
	'DataTable.ImportSheet clsEnvironmentVariables.MasterXLSPath,"Sheet1", "Global" 
	DataTable.ImportSheet clsEnvironmentVariables.MasterXLSPath,"Global", "Global" ''Teekam 20 Oct

    DataTable.GetSheet("Global").SetCurrentRow(1)
	EnvSheetPos=DataTable.GetSheetCount
	strEnvPos= EnvSheetPos
	clsEnvironmentVariables.EnvSheetPos = strEnvPos
	'strEnv=Replace(DataTable("Environment","Global")," ","")
  
	''Teekam 28 Oct  ''NOt required below environment  variable 	
	'strEnv=Replace(DataTable("Environment","Global")," ","")  

'Teekam -- 22-Feb - Commnted 
''	clsEnvironmentVariables.Environment = UCase(strEnv) 
	clsEnvironmentVariables.Environment =  AppURL   '' Storing URL value in environment variable
	 

	'clsEnvironmentVariables.TestCycle=DataTable("TestCycle",IntSheetNo-1)
   ' clsEnvironmentVariables.Build=DataTable("Build",IntSheetNo-1)
	clsEnvironmentVariables.TestCycle="1.0"
    clsEnvironmentVariables.Build="1.0"   

    DataTable.AddSheet "Master"	
	clsRepLoadScript.UploadRepository strControlPath
	clsRecoveryLoadScript.UploadRecovery strControlPath

'''---------------Teekam---------------------
''Teekam 18-Feb
'' Need to romve below code for chaning text recognization method
''---------------Teekam---------------------
	'---------------------------------------------------
	''Teekam 3 Aug 09
	''Teekam 22-Feb
''	Dim App 'As Application
''	Set App = CreateObject("QuickTest.Application")
''	App.Options.TextRecognitionOrder = "OCRThenAPI"
''	App.Options.TextRecognitionBlockType = "Single"
''	Set App= Nothing
	'---------------------------------------------------

	ReDim Preserve ResultFileName(DataTable.GetSheet("Global").GetRowCount)
	
	For iGrpCnt = 1 To DataTable.GetSheet("Global").GetRowCount
				DataTable.GetSheet("Global").SetCurrentRow(iGrpCnt)
 				strGroupName = Trim(DataTable("GroupName","Global"))

				 If Trim(UCase(DataTable("Execute","Global")))="Y" Then
					 ResultFileName (  iGrpCnt  -1 ) = strGroupName & "_RunManager.html"
					 
						 'Imports the master excel to the Master sheet of the run-time excel.
						strMasterFilePath = strControlPath & "Business_Scripts\TestCase_List\" & strProjectName & "\"  & Trim(DataTable("GroupName","Global")) & ".xls"					   
						strGroupName = Trim(DataTable("GroupName","Global"))

						'DataTable.ImportSheet  strMasterFilePath,"Sheet1", "Master"
						DataTable.ImportSheet  strMasterFilePath,"Global", "Master"  ''Teekam 20 Oct 

						ColCount=Datatable.GetSheet("Master").GetParameterCount
						intRowCountMaster = DataTable.GetSheet("Master").GetRowCount	
						 'Executes the Scenarios in order.						
						For iSheetCnt = 1 To intRowCountMaster   							
									DataTable.GetSheet("Master").SetCurrentRow(iSheetCnt)									
									If Trim(UCase(DataTable("Execute","Master")))="Y" Then				
											'If blnMultiUserExecution = True Then
													'If Trim(UCase(DataTable("ExecutedBy","Master")))=Trim(UCase(clsEnvironmentVariables.RunBy)) Then											
														Set clsDriverScript = New DriverScript

''----Teekam - 22-Feb----------------------------------------------------------------------
'''-- Not required in  EComm automation , no multiple iteration
''														If Datatable("Iteration","Master")="" OR Not IsNumeric(Datatable("Iteration","Master")) then
''																Datatable("Iteration","Master")=1
''														End If
''														Call clsDriverScript.DriverScripts( clsDatabase_Module,iSheetCnt, clsEnvironmentVariables, IntSheetNo, Datatable("Iteration","Master"))	
''-------------------------------------End Teekam 22-Feb ---------------------------------

														Call clsDriverScript.DriverScripts( clsDatabase_Module,iSheetCnt, clsEnvironmentVariables, IntSheetNo, 1 ) ''Teekam Iteration as 1 , 22-feb 
														Set clsDriverScript=Nothing	
													'End If
											'End If

									'''''----------------Teekam------------------------------
									''Teekam 18-Feb , need to  exit from this loop when find the execute as "y" and script execution from QC.												
									'''''----------------Teekam------------------------------

								End If
						Next
						Column_Count=Datatable.GetSheet("Master").GetParameterCount
'	
'						'Creates the summary report
						clsReport.Write_Summary_Header Column_Count-ColCount,clsEnvironmentVariables	
						clsReport.Add_Results_Summary Column_Count-ColCount,clsEnvironmentVariables  
'						'Deletes all the sheets created at the time of execution.
'						For iSheetCnt = DataTable.GetSheetCount to IntSheetNo + 1 step -1
'									Datatable.DeleteSheet iSheetCnt
'						Next 	
			  End If   			  
	Next

	clsDatabase_Module.Close_Database

''Teekam Start
	For iSheetCnt = DataTable.GetSheetCount to IntSheetNo + 1 step -1
			Datatable.DeleteSheet iSheetCnt
	Next 	
'''Teekam End


    'Desc: To handle QC integration
	'Close the excel file
	If UCase ( Environment.Value ( "TestName") ) = "RUNMANAGER" Then 
		''''-----------------Teekam 18-Feb ---------------------
		'' Code( below if conditiona and code inside if  condition ) is not more useful, while we will execute our script from QC for run time..
		''''-----------------Teekam---------------------------------
		If  blnQCUpdation Then
				clsQCIntegration_Module.CloseExcelFile clsEnvironmentVariables.UseExcelObject
				Set clsEnvironmentVariables.UseExcelObject  = Nothing
				'Upload all test results into QC from excel file
				clsQCIntegration_Module.UploadTestResultsInQCFromExcel  clsEnvironmentVariables.ServerNameQC,clsEnvironmentVariables.UserNameQC, clsEnvironmentVariables.PasswordQC,  clsEnvironmentVariables.DomainQC,   clsEnvironmentVariables.ProjectQC, clsEnvironmentVariables.TestSetPathQC , clsEnvironmentVariables.TestSetNameQC,clsEnvironmentVariables.TestResultExcelFile			
				clsQCIntegration_Module.UploadResultLogInQC  DriveName, strProjectName		   
		End If
	End If 


    'Invokes the Summary report at the end of the run.
	'To open the HTML report in a new browser. 
	Set WshShell = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fldr = fso.GetFolder(clsEnvironmentVariables.SummaryResultPath)
	Set flsc = fldr.Files
'	For Each fl In flsc
'		Rep_File_Path= clsEnvironmentVariables.SummaryResultPath & "\" & fl.Name
'	    WshShell.Run "iexplore " &  Rep_File_Path
'	Next 

	'TeekamQC  16  Sep 	
	If UCase ( Environment.Value ( "TestName") ) = "RUNMANAGER" Then	
		For i= 0 to  UBound(ResultFileName) 
			If  ResultFileName (i) <> ""Then
						Rep_File_Path= clsEnvironmentVariables.SummaryResultPath & "\" & ResultFileName (i)
						If fso.FileExists(TRim(Rep_File_Path)) Then
							WshShell.Run "iexplore " &  Rep_File_Path
						End If 	
			End If
		Next 
	End If
	Set fso = Nothing

	If UCase ( Environment.Value ( "TestName") ) <> "RUNMANAGER" Then

''''------------Teekam 18-Feb ----------------------------
'''' Need some modification and merging code with above line for performance improvement.
''''------------Teekam 18-Feb ----------------------------

			For i= 0 to  UBound(ResultFileName) 
					If  ResultFileName (i) <> ""Then
							strLogFilePath= clsEnvironmentVariables.TestResultPath & "\" & Replace ( ResultFileName ( i ) , "RunManager", "TestResultLog")
							Exit For 
					End If
			Next 
			
			''''Code for attaching log file with result  Teekam 20 _oct
			Set Attachments = QcUtil.CurrentRun.Attachments
			Set Attachment = Attachments.AddItem(Null)
			Attachment.FileName = strLogFilePath
			Attachment.Description = "HTML-Test Result Log "
			Attachment.Type = 1
			Attachment.Post 
			
			Set fso = CreateObject("Scripting.FileSystemObject")	
			If (fso.FileExists(strLogFilePath )) Then
				fso.deletefile(strLogFilePath )
			End If
			set fso = Nothing
	
	End If 
'' End of RunManager code
'''--------------------------------------------------------------------------------
''''-------------------------------------------------------------------------------

'********************************************************************************************************************************************************
'FUNCTION HEADER
'********************************************************************************************************************************************************
' Name:DownloadFrameworkOnMachine
' Description: This function is used to Download the framework and framewok related files from QC on local disk
'********************************************************************************************************************************************************
Function DownloadFrameworkOnMachine(ByVal DrivePath)

		Dim fso, QCPAYLESSProject
		
		QCPAYLESSProject ="Subject\PaylessEComm_RegressionAutomation\CRAFT"
		ProjectName = "PAYLESS_ECOMM"
    
		BusinessScriptsPath	="Business_Scripts"
		EnvironmentSetUpPath	="Environment_SetUp"
		FrameworkScriptsPath = "Framework_Scripts"

		'Connect to QC
		QCServerPath = "http://" & ServerName & "/qcbin/"

		Set QCConnection = QCUtil.QCConnection
	
		Set treeMgr  = QCConnection.TreeManager

		'Create the File System object to create folder
		Set fso = CreateObject("Scripting.FileSystemObject")

		TestCaseListPath = BusinessScriptsPath & "\TestCase_List\" & ProjectName
		
		' Download the Master excel file "PAYLESS_ECOMM_GoupMaster.xls"  and all master list files on local disk at  below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Buisness_Scripts\TestCase_List\PAYLESS_ECOMM\PAYLESS_ECOMM_Master.xls
		MasterTestCaseDownLoadPath = DrivePath & TestCaseListPath
		strFolderPath = QCPAYLESSProject & "\" & TestCaseListPath
		Call CreateFolderForProject (fso,MasterTestCaseDownLoadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,MasterTestCaseDownLoadPath)

		' Download the test case script excel files  and TestData file on local disk at  below path i.e.
		' C:\PaylessEComm_RegressionAutomation\CRAFT\Buisness_Scripts\TestCase_Scripts\PAYLESS_ECOMM
		TestCaseScriptPath = BusinessScriptsPath & "\TestCase_Scripts\" & ProjectName				
		strFolderPath = QCPAYLESSProject & "\" & TestCaseScriptPath
		TestCaseScriptDownLoadPath = DrivePath & TestCaseScriptPath
		Call CreateFolderForProject (fso,TestCaseScriptDownLoadPath)

		Set strParentFolder = treeMgr.NodeByPath(strFolderPath)		
        Set fc = strParentFolder.NewList()
		
		For Each sf In fc
				TestCaseScriptPath = BusinessScriptsPath & "\TestCase_Scripts\" & ProjectName & "\" & sf.Name 
				strFolderPath = QCPAYLESSProject & "\" & TestCaseScriptPath
				TestCaseScriptDownLoadPath = DrivePath & TestCaseScriptPath
				Call CreateFolderForProject (fso,TestCaseScriptDownLoadPath)
				Call DownloadFolderAttachment(treeMgr, strFolderPath,TestCaseScriptDownLoadPath)
		Next
			  
       'Download enviorment set up files at  below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Environment_SetUp\PAYLESS_ECOMM
		EnvironmentPath = EnvironmentSetUpPath & "\" & ProjectName
		TestEnvironmentDownloadPath = DrivePath & EnvironmentPath
		strFolderPath = QCPAYLESSProject & "\" & EnvironmentPath
		Call CreateFolderForProject (fso,TestEnvironmentDownloadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,TestEnvironmentDownloadPath)

		'Download files from Framework folder  at  below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Framework_Scripts\Driver_Script
		DriverScriptPath = FrameworkScriptsPath & "\Driver_Script"
		DriverScriptDownloadPath = DrivePath & DriverScriptPath
		strFolderPath = QCPAYLESSProject & "\" & DriverScriptPath
		Call CreateFolderForProject (fso,DriverScriptDownloadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,DriverScriptDownloadPath)

		'Download  Library folder and file at below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Framework_Scripts\Library_Script
		LibraryScriptPath = FrameworkScriptsPath & "\Library_Script"
		LibraryScriptDownloadPath = DrivePath & LibraryScriptPath
		strFolderPath = QCPAYLESSProject & "\" & LibraryScriptPath
		Call CreateFolderForProject (fso,LibraryScriptDownloadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,LibraryScriptDownloadPath)     

		'Download  Module folder and file at below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Framework_Scripts\Module_Script\PAYLESS_ECOMM
		ModuleScriptPath = FrameworkScriptsPath & "\Module_Script\" & ProjectName
		ModuleScriptDownloadPath = DrivePath & ModuleScriptPath
		strFolderPath = QCPAYLESSProject & "\" & ModuleScriptPath
		Call CreateFolderForProject (fso,ModuleScriptDownloadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,ModuleScriptDownloadPath)

		'Download  Object Repository and file at below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Framework_Scripts\ObjectRepository
		ObjectRepositoryPath = FrameworkScriptsPath & "\ObjectRepository"
		ObjectRepositoryDownloadPath = DrivePath & ObjectRepositoryPath
		strFolderPath = QCPAYLESSProject & "\" & ObjectRepositoryPath
		Call CreateFolderForProject (fso,ObjectRepositoryDownloadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,ObjectRepositoryDownloadPath)

		'Download  Recovery Scenario and file at below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Framework_Scripts\RecoveryScenarious
		RecoveryScenariousPath = FrameworkScriptsPath & "\RecoveryScenarios"
		RecoveryScenariousDownloadPath = DrivePath & RecoveryScenariousPath
		strFolderPath = QCPAYLESSProject & "\" & RecoveryScenariousPath
		Call CreateFolderForProject (fso,RecoveryScenariousDownloadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,RecoveryScenariousDownloadPath)

		'Download  Report Mangager and file at below path i.e.
		'C:\PaylessEComm_RegressionAutomation\CRAFT\Framework_Scripts\ReportManager
		ReportManagerPath = FrameworkScriptsPath & "\ReportManager"
		ReportManagerDownloadPath = DrivePath & ReportManagerPath
		strFolderPath = QCPAYLESSProject & "\" & ReportManagerPath
		Call CreateFolderForProject (fso,ReportManagerDownloadPath)
		Call DownloadFolderAttachment(treeMgr, strFolderPath,ReportManagerDownloadPath)

	    ReportPath = FrameworkScriptsPath & "\Reports\" & ProjectName
		RuntimeReportPath = DrivePath & ReportPath & "\Runtime_Reports"
		Call CreateFolderForProject (fso,RuntimeReportPath)

		ScreenShotPath = DrivePath & ReportPath & "\Screen_Shot"
		Call CreateFolderForProject (fso,ScreenShotPath)

		 SummaryPath = ReportPath & "\Summary"
		SummaryDownloadPath = DrivePath & SummaryPath
		Call CreateFolderForProject (fso,SummaryDownloadPath)
		strFolderPath = QCPAYLESSProject & "\" & SummaryPath
		'Added code to download the bmp files
		Call DownloadFolderAttachment(treeMgr, strFolderPath,SummaryDownloadPath)

		TestResultsLogPath = DrivePath & ReportPath & "\Test_Results_Log"
		Call CreateFolderForProject (fso,TestResultsLogPath)
		
End Function

'********************************************************************************************************************************************************
'FUNCTION HEADER
'********************************************************************************************************************************************************
' Name:DownloadMasterExcel
' Description: This function is used to download the Master Excel attachment from QC folder
' Input Parameter: None
'********************************************************************************************************************************************************
Function DownloadMasterExcel()

			QCPAYLESSProject ="Subject\"
			ProjectName = "PAYLESS_ECOMM"
			BusinessScriptsPath	="PaylessEComm_RegressionAutomation\CRAFT\Business_Scripts"

        'Connect to QC
			QCServerPath = "http://" & ServerName & "/qcbin/"
			Set QCConnection = QCUtil.QCConnection
            Set treeMgr  = QCConnection.TreeManager

		'Create the File System object to create folder
			Set fso = CreateObject("Scripting.FileSystemObject")

			TestCaseListPath = BusinessScriptsPath & "\TestCase_List\" & ProjectName
		' Download the Master excel file "PAYLESS_ECOMM_Master.xls" on local disk
			MasterTestCaseDownLoadPath = DriveName & DrivePath & TestCaseListPath
			strFolderPath = QCPAYLESSProject & "\" & TestCaseListPath

			Call DownloadFolderAttachment(treeMgr, strFolderPath,MasterTestCaseDownLoadPath)
			
End Function

'********************************************************************************************************************************************************
'FUNCTION HEADER
'********************************************************************************************************************************************************
' Name:DownloadFolderAttachment
' Description: This function is used to download the attachment from QC folder
' Input Parameter: TreeMgr, FolderName(Foldername from QC),DownloadPath(Local path to download the file)
'********************************************************************************************************************************************************
Function DownloadFolderAttachment(ByRef treeMgr ,ByVal FolderName,ByVal DownloadPath)
  
			Set attachFolder = treeMgr.NodeByPath(FolderName)
			' Call the FolderAttachment routine to download all the attachment
			' Get the Attachments 
			Set testAttachFact = attachFolder.Attachments
	
			' Get the list of attachments and go through 
			' the list, downloading one at a time. 
			Set attachList = testAttachFact.NewList("")

			For each tAttach In attachList
				Set attachemntstorage = tAttach.AttachmentStorage
						attachemntstorage.ClientPath = DownloadPath 

						QCFileName=tAttach.name(0)
						ActualFileName=tAttach.name(1)

						attachemntstorage.Load tAttach.name,True

						Set renfile = CreateObject("Scripting.FileSystemObject")

								If renFile.FolderExists(attachemntstorage.ClientPath) Then
									If  renFile.FileExists(attachemntstorage.ClientPath & "\" & ActualFileName) Then
										Set delfile = renFile.GetFile(attachemntstorage.ClientPath & "\" & ActualFileName)
										delfile.delete
									End If
									renFile.MoveFile attachemntstorage.ClientPath & "\" & QCFileName,attachemntstorage.ClientPath & "\" & ActualFileName 
							End If
			Next
							
End Function

'********************************************************************************************************************************************************
'FUNCTION HEADER
'********************************************************************************************************************************************************
' Name:VerifyUserMandatesOnDB
' Description: This function is used to crate a folder on local drive
' Input Parameter: fso(file system object), FolderPath(Folder to create)
'********************************************************************************************************************************************************
Function CreateFolderForProject(ByRef fso, ByVal strFolderPath)
	
	   If NOT fso.FolderExists(strFolderPath) Then
 			If  CreateFolderForProject (fso, fso.GetParentFolderName(strFolderPath)) Then
					Call fso.CreateFolder(strFolderPath)
					CreateFolderForProject = True
			End If
	   Else
				CreateFolderForProject = True
	  End If 

End Function


'********************************************************************************************************************************************************
'FUNCTION HEADER
'********************************************************************************************************************************************************
' Name:Group_Test_Case_Modify
' Description: This procedure is used to mark "N"  in all value in execute  and "Y" with  passed named 
' Input Parameter: strFileName and strTestName : file name and only value mark as "Y"
'********************************************************************************************************************************************************
Sub Group_Test_Case_Modify ( strFileName, strTestName,Col_Name )
	
	strFileLocation = "C:\PaylessEComm_RegressionAutomation\CRAFT\Business_Scripts\TestCase_List\PAYLESS_ECOMM"
	
	''For updating Execute column all value as "N"
	DataTable.ImportSheet strFileLocation & "\" & strFileName & ".xls"  ,1 ,1 
	rowcount = DataTable.GetSheet("Global").GetRowCount 

	For i=1 To rowcount
		DataTable.SetCurrentRow( i )		
		Execute_Flag = DataTable.Value("Execute","Global")
		ColName = DataTable.Value(Col_Name,"Global")

		If UCase(Execute_Flag) ="Y" Then        
			DataTable.Value("Execute","Global")="N"			
		End If   
	Next

	For i=1 To rowcount
		DataTable.SetCurrentRow(i)
		Execute_Flag = DataTable.Value("Execute","Global")
		ColValue  = DataTable.Value(Col_Name,"Global")
		If UCase(ColValue)=UCase(strTestName) Then
			DataTable.Value("Execute","Global")="Y"
			Exit For				   
		End If   
	Next
	DataTable.ExportSheet strFileLocation & "\" & strFileName & ".xls"  ,1

End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++