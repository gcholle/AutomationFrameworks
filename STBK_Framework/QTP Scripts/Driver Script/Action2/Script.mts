''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Script Name          Main Script
'    Author             CAT Team
'    Date Created         Mar-09
'    Purpose            Main Script for STBK
'
'   Changes           Date     By          Description 
'   ------------------------------------------------------------------------------------------------------------------------------------
 
'`````````````````````````````````````````````````````
' Variable    Declaration
'`````````````````````````````````````````````````````
Option Explicit
 
Dim objExcel_Act,stTimr,endTimr,sfailCnt,sfilename,z,Declaration_Ctr
Dim j, functionStatus, reportStatus, TestCaseIdNext, ExitScenario,LoginScenario
Dim  row,col,FieldLength,CurrentRow,TestCaseId ,ScreenName,Action,Comment
Dim ScreenMapPath,TestDataPath,DetailedSheetPath
Dim numberOfFields,step_status,SearchedRow
Dim hr,min,sec,dtTemp,szMonth,szDay,szYear,MyTime
Public DataRow,ObjGLSheet, ObjDataSheet,ObjDetailSheet,Con_Screenmap
Public SystemName, TestIdentifier,step_result,restoreLogin
Dim ExceptionHandler, TestCaseIdCurrItr, TestCaseIdNextItr,doExit,LoginFail,Status,fn,fv,ReportName
Dim sumPrint,y,ApplicationName,TestData_ctr,r_count,DR_Count,ExceptionTestCaseID
Dim Update_TD,SummaryReportFile, TestLabTestSetPath, TestLabTestSetName,TD_Path,tdPathArray,tdPathSubscript,i,qtApp,ResultFolderPath
 
''Reading the ReportName
ReportName = parameter("ReportName")
ResultFolderPath = parameter("ResultFolderPath")
 
'Setting Current Row
'CurrentRow=DataTable.GetCurrentRow
DataTable.SetCurrentRow(1)
 
'Path where TestData is stored
If parameter("DataSheetName") <> "" Then
 TestDataPath =  PathFinder.Locate("..\..\TestData\"&parameter("DataSheetName")&".xls")
End If
 

If ReportName = "" Then
 ReportName = "No Name" 
End If
 
'Code to create excel object for reporting 
reportStatus=addsheets_Report (ResultFolderPath,ReportName,sfilename)
Set objExcel_Act = CreateObject("Excel.Application")
objExcel_Act.Workbooks.Open sfilename
objExcel_Act.Visible = true
sumPrint = "NO"
 
''CAdodb connection to connect to Screenmap.mdb
' Set Con_Screenmap = CreateObject("Adodb.connection")
' Con_Screenmap.Provider = "microsoft.jet.oledb.4.0"
' Con_Screenmap.Properties("Data Source").Value = parameter("ScreenMapPath")
' Con_Screenmap.open
 
 stTimr = Hour(now)&":"&Minute(now)&":"&Second(now)  '' it will store Start time
 ''To check the testdata file retreive from the Driversheet actually exist 
If TestDataPath="" Then
  ScreenName = ""
  Action = ""
  TestCaseID = "" 
  fn = "TestData Sheet"
  Status =  "FAILED"
  Comment = "Testdata file mentioned in the DriverSheet does not exist."
  step_status=appendValidationSheet(objExcel_Act, ScreenName,ApplicationName, Action, TestCaseID,fn,fv,Status, Comment, sfailCnt )
  sfailCnt = sfailCnt + 1
  TestCaseID = "TestData sheet not found"
  appendSummarySheet objExcel_Act,TestCaseID, stTimr,endTimr,sfailCnt
  stTimr = endTimr
  endTimr=0
  sfailCnt = 0
Else
 
'Importing Test Data Sheet
DataTable.AddSheet "DataSheet"           
DataTable.ImportSheet TestDataPath,"Sheet1","DataSheet"
 
'Setting object of Datatable
Set ObjDataSheet = DataTable.GetSheet("DataSheet")
 
''Code to check and connect if session is disconnected
'functionStatus= SessionConnect(TestCaseID,objExcel_Act,ExitScenario, Comment)
 

'To validate if rows in TDS are blank
TestData_ctr = 0
ObjDataSheet.SetCurrentRow (1)
TestCaseId = trim(DataTable.Value("TestCaseID","DataSheet"))
Do  While TestCaseId <> ""
 TestCaseId = trim(DataTable.Value("TestCaseID","DataSheet"))
 If TestCaseId = "" Then
  Exit Do
 End If
 If TestData_ctr = ObjDataSheet.GetRowCount Then
  Exit Do
 End If
  TestData_ctr = TestData_ctr + 1
  ObjDataSheet.SetNextRow
Loop
 
'Main Loop which will iterate through the rows in excel file
For DataRow =1 to TestData_ctr
 
    ObjDataSheet.SetCurrentRow (DataRow) 
 TestCaseId = trim(DataTable.Value("TestCaseID","DataSheet"))
 If TestCaseId <>"" Then
 
 ScreenName = Trim(DataTable.Value("ScreenName","DataSheet"))
 Action= Trim(DataTable.Value("Action","DataSheet"))
 
'Derive the number of  fields inserted for a particular action
 For j=1to 256
  If (DataTable.Value("FieldName"& j,"DataSheet") = "") then      
   numberOfFields = j-1
   Exit for
  Else
   numberOfFields = ""
  End If   
 Next
 If  numberOfFields = 0 Then
  numberOfFields = ""
 End If
'Getting Application Name
 ApplicationName = Trim(DataTable.Value("ApplicationName","DataSheet"))

'********************************************************************************************************************************************************************************************************
'Seperatation of TOPS & COMET Application Sections
'********************************************************************************************************************************************************************************************************
 
If  ApplicationName ="TOPS" or  ApplicationName ="STBK" Then
If Environment.Value("Declaration_Ctr") ="0" Then
		'*******************************************
	'Define ScreenMap Path
	'*******************************************
	If  ApplicationName ="TOPS" Then
		ScreenMapPath =  PathFinder.Locate("..\..\Object Repositories\TOPS\TOPSObjectRepository.mdb")
	Else
		ScreenMapPath =  PathFinder.Locate("..\..\Object Repositories\STBK\STBKObjectRepository.mdb")
	End If
	
'	'*******************************************
'	'Import Login Credentials Sheet
'	'*******************************************
'	DataTable.AddSheet "Login_credentials"  '' for storing the login credentials
'	''''create the columns for the different fields in the login credentials
'	For z=1 to 11
'		functionStatus=DataTable.AddSheet ("Login_credentials").AddParameter("FieldName"&z, "")	  
'		functionStatus=DataTable.AddSheet ("Login_credentials").AddParameter("FieldValue"&z, "")	  
'	Next
'		functionStatus=DataTable.AddSheet ("Login_credentials").AddParameter("numberOfFields", "")	
		
	'***********************************************************
	'CAdodb connection to connect to Screenmap.mdb
	'***********************************************************
	 Set Con_Screenmap = CreateObject("Adodb.connection")
	 Con_Screenmap.Provider = "microsoft.jet.oledb.4.0"
	 Con_Screenmap.Properties("Data Source").Value = ScreenMapPath
	 Con_Screenmap.open
	
	 Environment.Value("Declaration_Ctr") = "1"
End If

'**************************************************************
 'Code to check and connect if session is disconnected
 '*************************************************************
functionStatus= SessionConnect(TestCaseID,objExcel_Act,ExitScenario, Comment)

'#########################################################################################################3
 
  functionStatus= DataTable.GetSheet("DataSheet").SetPrevRow
   'If  (DataTable.GetSheet("DataSheet").GetCurrentRow <> DataTable.GetSheet("DataSheet").GetRowCount) Then
   If  (DataTable.GetSheet("DataSheet").GetCurrentRow <> TestData_ctr) Then
 TestCaseIdNext = Trim(DataTable.Value("TestCaseID","DataSheet"))
    If  TestCaseIdNext <> TestCaseID Then
  ''code for restoring if in normal scenario
   If (Ucase(trim(Action)) <> "LOGIN") or (Ucase(trim(Action)) <> "LOGOFF") Then 
    LoginScenario="N"
   else
   LoginScenario="Y"
   End If
    End If
 End if
 
 functionStatus=DataTable.GetSheet("DataSheet").SetNextRow
 Action= Trim(DataTable.Value("Action","DataSheet"))
 ScreenName = DataTable.Value("ScreenName","DataSheet")
 
'Code to handle navigation to HOME screen irrespective of it's current location while in between the script
 If ( UCASE(Action) = "INPUT"  AND UCASE(ScreenName) = "HOME" AND UCASE(DataTable.Value("FieldName1","DataSheet")) = "CONTROL LINE") Then
  i=1
  While ValidateScreen(Con_Screenmap,"Home",Comment)<>micPass and i = 5
   TeWindow("TeWindow").TEScreen("TeScreen").SendKey TE_CLEAR
   WaitTillBusy
     i= i+1
  Wend
  i = ""
 End If

 'Code to handle Clear screen before Main Screen screen(Only for STBK)
 If ( UCASE(Action) = "INPUT"  AND UCASE(ScreenName) = "MAIN_MENU") Then
  TeWindow("TeWindow").TEScreen("TeScreen").SendKey TE_CLEAR
  WaitTillBusy
  TeWindow("TeWindow").TeScreen("TeScreen").SendKey TE_HOME
	For i = 1 to 15
		TeWindow("TeWindow").TeScreen("TeScreen").SendKey TE_ERASE_EOF
		WaitTillBusy
		TeWindow("TeWindow").TeScreen("TeScreen").SendKey TE_TAB
	Next

 End If
 'Code to handle Clear screen before Process_Initator screen(Only for STBK)
 If UCASE(ScreenName) = "PROCESS_INITIATOR" Then
	TeWindow("TeWindow").TEScreen("TeScreen").SendKey TE_CLEAR
	WaitTillBusy
 End If
'Code to handle navigation of EDS screen specifically from EDS_10 to EDS_5 and from EDS_11 to EDS_6 by pressing PF3 key
 If  UCASE(ScreenName) = "EDS_5"  or UCASE(ScreenName) = "EDS_6" Then
   functionStatus=DataTable.GetSheet("DataSheet").SetPrevRow
   ScreenName = DataTable.Value("ScreenName","DataSheet")
   If  UCASE(ScreenName) = "EDS_10" or UCASE(ScreenName) = "EDS_11" Then
   TeWindow("TeWindow").TEScreen("TeScreen").SendKey TE_PF3
   WaitTillBusy
   ElseIf  UCASE(ScreenName) = "EDS_5_5" or UCASE(ScreenName) = "EDS_6_5" Then
   TeWindow("TeWindow").TEScreen("TeScreen").SendKey TE_PF3
   WaitTillBusy
  End If
  functionStatus=DataTable.GetSheet("DataSheet").SetNextRow
  ScreenName = DataTable.Value("ScreenName","DataSheet")
 End If
 ' Code to check if Employee Policy Pick Screen & Patient Pick Screen is appearing or not
Select Case Ucase(ScreenName)
	Case "EMPLOYEE_POLICY_PICK"
		If (ValidateScreen(Con_Screenmap,"Patient_Pick",Comment) = micPass or ValidateScreen(Con_Screenmap,"Provider_Pick",Comment) = micPass or ValidateScreen(Con_Screenmap,"HCFA_Prov_Pat_Info",Comment) = micPass or ValidateScreen(Con_Screenmap,"UB92_Prov_Pat_Info",Comment) = micPass) Then
			Environment.Value ("Emp_Policy_Pick") = "N"
			Environment.Value ("Provider_Pick") = ""
			Environment.Value ("Patient_Pick") = ""
		End If
	Case  "PATIENT_PICK"
		If (ValidateScreen(Con_Screenmap,"Provider_Pick",Comment) = micPass or ValidateScreen(Con_Screenmap,"HCFA_Prov_Pat_Info",Comment) = micPass or ValidateScreen(Con_Screenmap,"UB92_Prov_Pat_Info",Comment) = micPass) Then
			Environment.Value ("Emp_Policy_Pick") = ""
			Environment.Value ("Provider_Pick") = ""
			Environment.Value ("Patient_Pick") = "N"
		End If
	Case "PROVIDER_PICK"
		If (ValidateScreen(Con_Screenmap,"HCFA_Prov_Pat_Info",Comment) = micPass or ValidateScreen(Con_Screenmap,"UB92_Prov_Pat_Info",Comment) = micPass) Then
			Environment.Value ("Emp_Policy_Pick") = ""
			Environment.Value ("Patient_Pick") = ""
			Environment.Value ("Provider_Pick") = "N"
		End If
	Case Else
		Environment.Value ("Emp_Policy_Pick") = ""
		Environment.Value ("Patient_Pick") = ""
		Environment.Value ("Provider_Pick") = ""
End Select

' 
'If UCASE(ScreenName) = "EMPLOYEE_POLICY_PICK" AND (ValidateScreen(Con_Screenmap,"Patient_Pick",Comment) = micPass or ValidateScreen(Con_Screenmap,"Provider_Pick",Comment) = micPass) Then
'	Environment.Value ("Emp_Policy_Pick") = "N"
'	Datatable.GetSheet("DataSheet").SetNextRow
'	Action = DataTable.Value("Action","DataSheet")
'	ScreenName = DataTable.Value("ScreenName","DataSheet")
'ElseIf (UCASE(ScreenName) = "PATIENT_PICK" AND ValidateScreen(Con_Screenmap,"Provider_Pick",Comment) = micPass) Then
'	Environment.Value ("Patient_Pick") = "N"
'	Datatable.GetSheet("DataSheet").SetNextRow
'	Action = DataTable.Value("Action","DataSheet")
'	ScreenName = DataTable.Value("ScreenName","DataSheet")
'Else
'	Environment.Value ("Emp_Policy_Pick") = ""
'	Environment.Value ("Patient_Pick") = ""
'End If
 
'#########################################################################################################
 
 Select Case ucase(Action)
  Case ucase("Login")
   step_status=  loginToTOPS (Con_Screenmap,numberOfFields, Action, TestCaseID,objExcel_Act, ExitScenario,LoginScenario,ExceptionHandler,LoginFail,sfailCnt,ApplicationName)
  Case ucase("Input")
            step_status= InputFunction(Con_Screenmap,SearchedRow,numberOfFields,ScreenName,Action,TestCaseID,objExcel_Act,ExitScenario,sfailCnt,ApplicationName)
  Case ucase("Search")
            step_status= Search(Con_Screenmap,numberOfFields,ScreenName,Action,TestCaseID,objExcel_Act,SearchedRow,ExitScenario,sfailCnt,ApplicationName)
  Case ucase("Validate")
            step_status= OutputFunction(Con_Screenmap,numberOfFields,ScreenName,Action,TestCaseID,objExcel_Act,ExitScenario,sfailCnt,ApplicationName)
  Case ucase("LogOff")
            step_status= LOGOFF(Con_Screenmap,ScreenName,Action,TestCaseID,objExcel_Act,ExitScenario,sfailCnt,ApplicationName)
  Case ucase("LogOff_CC")
            step_status= LogOff_CC(Con_Screenmap,ScreenName, Action, objExcel_Act,TestCaseID, ExitScenario,sfailCnt,ApplicationName)
  Case ucase("Login_CC")
   step_status = Login_CC(Con_Screenmap,TestCaseID, Action,objExcel_Act, Comment, Row, Col, FieldLength, sfailCnt,ExceptionHandler,ExitScenario,ApplicationName)
  Case ucase("LogOff_STBK")
   step_status= LogOff_STBK(Con_Screenmap,ScreenName,Action,TestCaseID,objExcel_Act,ExitScenario,sfailCnt,ApplicationName)
 End Select
 
'Code for skipping a scenario if exit scenario is YES
    call Exception_Handling(TestCaseID,Action,objExcel_Act,ExitScenario,Comment,ExceptionHandler,sfailCnt,ApplicationName)
    If ExitScenario = "YES" Then
  'If  (DataTable.GetSheet("DataSheet").GetCurrentRow <> DataTable.GetSheet("DataSheet").GetRowCount) Then
  If  (DataTable.GetSheet("DataSheet").GetCurrentRow <> TestData_ctr) Then
  TestCaseIDCurrItr= DataTable.Value("TestCaseID","DataSheet")
  Do 
   Action= DataTable.Value("Action","DataSheet")
   If (ucase(Action) = "LOGOFF") or (ucase(Action) = "LOGOFF_CC") or  (ucase(Action) = "LOGOFF_STBK") Then
    Exit Do
   End If 
   functionStatus=DataTable.GetSheet("DataSheet").SetNextRow
   TestCaseIdNextItr= DataTable.Value("TestCaseID","DataSheet")
   DataRow = DataTable.GetSheet("DataSheet").GetCurrentRow 
   If DataRow =  TestData_ctr Then
                doExit = "YES"
    Exit Do
   else
    doExit = "NO" 
   End If
  Loop while TestCaseIdNextItr = TestCaseIDCurrItr
  If ExceptionHandler <> "YES" Then
   sumPrint = "NO"
  End If
 
 Action= Trim(DataTable.Value("Action","DataSheet"))
 If ucase(Action) = "LOGOFF"  Then
   step_status= LOGOFF(Con_Screenmap,ScreenName,Action,TestCaseID,objExcel_Act,ExitScenario,sfailCnt,ApplicationName) 
   DataRow = DataRow + 1
   elseIf ucase(Action) = "LOGOFF_STBK"  Then
   step_status= LogOff_STBK(Con_Screenmap,ScreenName,Action,TestCaseID,objExcel_Act,ExitScenario,sfailCnt,ApplicationName) 
   DataRow = DataRow + 1
   elseIf ucase(Action) = "LOGOFF_CC" Then
   Action= DataTable.Value("Action","DataSheet")
   If Environment.Value("Login_CC_Flag") = "Y" Then
     Action= "Login_CC"
     step_status = Login_CC(Con_Screenmap,TestCaseID, Action,objExcel_Act, Comment, Row, Col, FieldLength, sfailCnt, ExceptionHandler,ExitScenario,ApplicationName)
  End If
   'step_status = Login_CC(Con_Screenmap,TestCaseID, Action,objExcel_Act, Comment, Row, Col, FieldLength, sfailCnt,ExceptionHandler,ExitScenario,ApplicationName)
   'Action= "Login_CC"
   step_status= LogOff_CC(Con_Screenmap,ScreenName, Action, objExcel_Act,TestCaseID, ExitScenario,sfailCnt,ApplicationName)
   DataRow = DataRow + 1
 End If 
'''=====================================
  
  If (ucase(DataTable.Value("Action","DataSheet")) <> "LOGIN") OR (UCASE(DataTable.Value("Action","DataSheet")) = "LOGOFF") Then
   If  TestCaseIdNextItr <> TestCaseIDCurrItr Then
   If ExceptionHandler = "YES" Then'    
'     ScreenName = DataTable.Value("ScreenName","DataSheet")
'     Action= DataTable.Value("Action","DataSheet")
'                If Environment.Value("Login_CC_Flag") = "Y" Then
'     step_status = Login_CC(Con_Screenmap,TestCaseID, Action,objExcel_Act, Comment, Row, Col, FieldLength, sfailCnt, ExceptionHandler,ExitScenario,ApplicationName)
'    End If
'    If ucase(LoginFail) <>"YES" Then
'       step_result = loginToTOPS (Con_Screenmap,numberOfFields, Action, TestCaseID,objExcel_Act, ExitScenario,LoginScenario,ExceptionHandler,LoginFail,sfailCnt,ApplicationName)
'    End If     
'   Else 
'    step_result = navigateToHome(Con_Screenmap)
	Exit For
   End If
        End If
   End If
 
''=======================================
  If  doExit = "NO" Then
   DataRow = DataRow-1
  End If
  
  End If
 End if
 

'Code to create Summary Report

datatable.GetSheet("DataSheet").SetNextRow
TestCaseIdNext = DataTable.Value("TestCaseID","DataSheet")
If trim(TestCaseIDNext) <> trim(TestCaseID) OR DataRow = TestData_ctr Then
		If LoginFail = "YES" Then
			  TestCaseID = TestCaseIDCurrItr
			  If DataRow =  TestData_ctr Then
				   doExit = "YES"
				   Comment = TestCaseIDCurrItr& " is not Executed as same previous Login credentials"
				   Status = "FAILED"
				   TestCaseID = TestCaseIDCurrItr
					Call appendValidationSheet(objExcel_Act, ScreenName, ApplicationName, Action, TestCaseID,fn,fv,Status, Comment, sfailCnt )
				End If
			End If
'		 If  doExit = "NO" Then
'			DataRow = DataRow-1
'		End If 
       appendSummarySheet objExcel_Act,TestCaseID, stTimr,endTimr,sfailCnt
       stTimr = endTimr
       endTimr=0
       sfailCnt = 0
End If



Else
   fn = ""
   fv = ""
   Status = "FAILED"
   Comment = "COMET application not yet integrated with the framework"
   y=appendValidationSheet(objExcel_Act, ScreenName,ApplicationName, Action, TestCaseID,fn,fv,Status, Comment, sfailCnt )
   If sumPrint <> "YES" Then
    appendSummarySheet objExcel_Act,TestCaseID, stTimr,endTimr,sfailCnt
    sumPrint = "YES"
   End If
   
  End If
 End If


 
Next
 
DataTable.DeleteSheet "DataSheet"
Set Con_Screenmap = nothing
 End If
 
    '**********************************
  'Code for TD Integration
  '**********************************
  DataTable.GetSheet("DriverSheet").SetCurrentRow(Environment.Value("DriverSheetCurrentRow"))
  Update_TD = DataTable.Value("Update_TD","DriverSheet")
  If ucase(Update_TD) = "YES" Then
   On error resume next
   TD_Path = DataTable.Value("TD_Path","DriverSheet")
   tdPathArray = Split(TD_Path,"\")
   tdPathSubscript = ubound(tdPathArray)
   TestLabTestSetName = tdPathArray(tdPathSubscript)
   For i =0 to tdPathSubscript -1
    If i = 0 Then
     TestLabTestSetPath =  TestLabTestSetPath & tdPathArray(i)
    Else
     TestLabTestSetPath =  TestLabTestSetPath & "\" & tdPathArray(i)   
    End If 
   Next
   
   'Report_Name = DataTable.Value("Report_Name","DriverSheet")
   If QCUtil.IsConnected Then
    Set qtApp = CreateObject("QuickTest.Application")  
    qtApp.TDConnection.Disconnect ' Disconnect from Quality Center 
   End If
   SummaryReportFile= sfilename
   UPDATE_TD_RESULTS objExcel_Act,SummaryReportFile, TestLabTestSetPath, TestLabTestSetName
  End If
  
  Set qtApp = nothing
  objExcel_Act.quit
  Set objExcel_Act = nothing
  'Set Con_Screenmap = nothing
  Set ObjDataSheet = nothing











