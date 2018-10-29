'***********************************************************************************************************************************************************
'Script Name: Driver Script for STBKApplication
'Team: CAT
'Creation Date: March ' 09
'Input Parameters:
'Output Parameters:
'Description: This script perfroms the follwoing activities:
'								- Importing Driver Sheet
'						    	- Local Result Folder Path
'						    	- Define ScreenMap Path
'						    	- Login Credentials
'						    	- Calling Module Scripts
'						    	- TD Integration
'***********************************************************************************************************************************************************
Option Explicit
'*******************************************
'Variables Declaration and Initialization 
'*******************************************
Dim ScreenMapPath,driverSheet, ResultFolderPath
Dim ExecuteModule, DataSheetName, ReportName,DriverDataRow,wdProcessId, temp_ctr
Dim AllProcess,Process

'*******************************************
'Importing Driver Sheet
'*******************************************
driverSheet=  PathFinder.Locate("..\..\Scheduler\DriverSheet.xls")
DataTable.AddSheet "DriverSheet"
DataTable.ImportSheet driverSheet,"DriverSheet","DriverSheet"

'*******************************************
'Local Result Folder Path
'*******************************************
ResultFolderPath = PathFinder.Locate("..\..\Results")



'*******************************************
'Performance Tuning script
'*******************************************
ExecuteFile PathFinder.Locate("..\..\Environment\TestSetting.vbs")
ExecuteFile PathFinder.Locate("..\..\Environment\QTP Settings.vbs")




'*******************************************
'Calling Module Scripts
'*******************************************
For DriverDataRow =1 to DataTable.GetSheet("DriverSheet").GetRowCount
	'**************************************************
	'Setting Counter to execute declaration only once for TOPS
	'**************************************************
	Environment.Value("Declaration_Ctr") ="0"

	Environment.Value("DriverSheetCurrentRow") = DriverDataRow
	DataTable.GetSheet("DriverSheet").SetCurrentRow(DriverDataRow)
	ExecuteModule = DataTable.Value("Execute","DriverSheet") 
	If Ucase(ExecuteModule) = "YES" Then
		Environment.Value("CaptureAllScreens") = DataTable.Value("Screen_Capture","DriverSheet")
		ReportName = DataTable.Value("Report_Name","DriverSheet")
		DataSheetName = DataTable.Value("Data_Sheet_Name","DriverSheet")
'		ExecuteFile PathFinder.Locate("..\..\Function Library\TOPS\TOPSFunctionLibrary.vbs")

		'Creation of Result Folder
		If (ucase(Environment.Value("CaptureAllScreens"))="ALWAYS") OR (ucase(Environment.Value("CaptureAllScreens"))="ON ERROR") Then
			Dim fso, f,FolderName,ErrorFolder
			Set fso = CreateObject("Scripting.FileSystemObject")
			FolderName=(Month(Date))&"_"&Day(Date)&"_"&Year(Date)&"_"&Hour(now)&"_"&Minute(now)&"_"&Second(now)
			Set f = fso.CreateFolder(ResultFolderPath&"\"&FolderName)
			Environment.Value("ScreenShotFolderPath") = f.Path
            Set fso=nothing
			Set f=nothing
		End If
				'RunAction "Main Script [Main Script]", oneIteration, ReportName, ScreenMapPath, DataSheetName,ResultFolderPath
        RunAction "Main Script", oneIteration, ReportName, DataSheetName,ResultFolderPath
	End If
	
Next
'Close TE WIndow

Set AllProcess = getobject("winmgmts:") 'create object 
For Each Process In AllProcess.InstancesOf("Win32_process") 'Get all the processes running in your PC 
	If (Instr (Ucase(Process.Name),"PCSCM.EXE") = 1) or (Instr (Ucase(Process.Name),"PCSWS.EXE") = 1)  Then 'Made all uppercase to remove ambiguity. 
		If  (Instr (Ucase(Process.Name),"PCSCM.EXE") = 1)  Then
			SystemUtil.CloseProcessByName "pcscm.exe"
		Else
			SystemUtil.CloseProcessByName "pcsws.exe"
	End If
	End If 
Next
Set AllProcess = Nothing

