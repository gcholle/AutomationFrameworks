'#####################################################################################################################
'Script Description				: Driver Script to trigger the testcase execution
'Test Tool/Version				: HP Quick Test Professional 11.53 and above
'Test Tool Settings				: N.A.
'Application Automated			: Health Connector
'Author							: Cigniti
'Date Created					: 29/05/2014
'Modified by : Ram, Rpeorting Craft event.. (24-June-2015)
'#####################################################################################################################
Option Explicit	'Forcing Variable declarations

'Declare required variables
Dim gobjFso
Dim gstrProjectName
Dim gstrDatatableName
Dim gstrBusinessFlowSheet, gstrCheckPointSheet
Dim gstrResultSheet, gstrReportedEventSheet, gstrCurrentScenario
Dim gstrCurrentTestCase, gstrIterationMode, gintStartIteration
Dim gintEndIteration, gintCurrentIteration, gintCurrentBusinessFlowRow
Dim gintCurrentTestDataRow, gintCurrentReportedEventRow, gintCurrentFlowNumber
Dim garrCurrentFlowData, gstrCurrentKeyword, gintGroupIterations
Dim gintCurrentGroupIteration, garrComponentGroup(), gintGroupedComponents
Dim gintCurrentComponent, gobjHashTable

'Initialize basic configuration settings from CRAFT.ini file
gstrBusinessFlowSheet =  CRAFT_GetConfig("BusinessFlowSheet")
gstrCheckPointSheet = CRAFT_GetConfig("CheckPointSheet")
gstrResultSheet = CRAFT_GetConfig("ResultSheet")
gstrReportedEventSheet = CRAFT_GetConfig("ReportedEventSheet")
Environment.Value("CheckpointSheet") = gstrCheckpointSheet
Environment.Value("ReportedEventSheet") = gstrReportedEventSheet
Environment.Value("ResultSheet") = gstrResultSheet
Environment.Value("TakeScreenshotFailedStep") = _
							CBool(CRAFT_GetConfig("TakeScreenshotFailedStep"))
Environment.Value("TakeScreenshotPassedStep") = _
							CBool(CRAFT_GetConfig("TakeScreenshotPassedStep"))
'Options are NextIteration, NextTestCase, NextStep, Stop, Dialog
Environment.Value("OnError") = CRAFT_GetConfig("OnError")	
If CBool(CRAFT_GetConfig("DebugMode")) Then
	'Turn off error handling to enable debugging
	Environment.Value("OnError") = "NextStep"	
End If
Environment.Value("ReportsTheme") = CRAFT_GetConfig("ReportsTheme")
Environment.Value("ResultPath") = Pathfinder.Locate("Results")
Environment.Value("OverallStatus") = ""
Environment.Value("RunIndividualComponent") = False
Environment.Value("TestCase_ExecutionTime") = 0

'Setup appropriate parameters for the Current Test Case Execution (passed from the initialization script)
gstrCurrentScenario = TestArgs("CurrentScenario")
Environment.Value("CurrentScenario") = gstrCurrentScenario
gstrCurrentTestCase = TestArgs("CurrentTestCase")
Environment.Value("CurrentTestCase") = gstrCurrentTestCase
Environment.Value("TimeStamp") = TestArgs("TimeStamp")
gstrIterationMode = TestArgs("IterationMode")
gstrDatatableName = gstrCurrentScenario

'Import required sheets from Datatable
Set gobjFso = CreateObject("Scripting.FileSystemObject")
If Not gobjFso.FileExists(Pathfinder.Locate("Datatables\") & gstrDatatableName & ".xls") Then
	Reporter.ReportEvent micFail,"Error",_
						"Datatable not found for the specified Scenario!"
	ExitRun
End If

CRAFT_ImportSheet Pathfinder.Locate("Datatables\") & gstrDatatableName & ".xls",_
													gstrBusinessFlowSheet
CRAFT_ImportSheet Pathfinder.Locate("Datatables\") & gstrDatatableName & ".xls",_
													gstrReportedEventSheet
If gobjFso.FileExists(Environment.Value("ResultPath") & "\" &_
		Environment.Value("TimeStamp") & "\Excel Results\Summary.xls") Then
	CRAFT_ImportSheet Environment.Value("ResultPath") & "\" &_
		Environment.Value("TimeStamp") & "\Excel Results\Summary.xls",_
														gstrResultSheet
Else
	CRAFT_ImportSheet Pathfinder.Locate("Datatables\") & gstrDatatableName & ".xls",_
																gstrResultSheet
End If
Set gobjFso = Nothing

'Setup the test case iterations
Select Case gstrIterationMode
	Case "oneIteration"
		gintStartIteration = 1
		gintEndIteration = 1
		gintCurrentIteration = 1
	Case "rngIterations"
		gintStartIteration = TestArgs("StartIteration")
		gintEndIteration = TestArgs("EndIteration")
		If gintStartIteration = "" then
			gintStartIteration = 1
		End if
		If gintEndIteration = "" then
			gintEndIteration = 1
		End if
		gintCurrentIteration = gintStartIteration
	Case "rngAll"
		gintStartIteration = 1
		gintEndIteration = 65535
		gintCurrentIteration = 1
End Select

'Execute all iterations of Current Test Case
Set gobjHashTable = CreateObject("Scripting.Dictionary")

gintCurrentBusinessFlowRow = _
		CRAFT_SetBusinessFlowRow(gstrCurrentTestCase, gstrBusinessFlowSheet)
Do while CInt(gintCurrentIteration) <= CInt(gintEndIteration)
	Environment.Value("CurrentIteration") = gintCurrentIteration
	
	CRAFT_ReportEvent gstrReportedEventSheet, "Start",_
						"Iteration" & gintCurrentIteration & " started", "Completed"
	Environment.Value("Iteration_StartTime") = Now()
	Environment.Value("ExitIteration") = False
	Environment.Value("StopExecution") = False
	
	gintCurrentFlowNumber = 1
	gintGroupedComponents = 0
	garrCurrentFlowData = _
				Split(DataTable.Value("Keyword_1",gstrBusinessFlowSheet),",")
	gstrCurrentKeyword = garrCurrentFlowData(0)
	Do until gstrCurrentKeyword = ""
		If UBound(garrCurrentFlowData) = 0 Then
			gintGroupIterations = 1
		Else
			gintGroupIterations = garrCurrentFlowData(1)
		End If
		
		gintGroupedComponents = gintGroupedComponents + 1
		Redim Preserve garrComponentGroup(gintGroupedComponents)
		garrComponentGroup(gintGroupedComponents - 1) = gstrCurrentKeyword
		
		If (gintGroupIterations > 0) Then	'Reached the end of a group (a group may comprise only one keyword also)
			For gintCurrentGroupIteration = 1 To gintGroupIterations	'Execute all group iterations specified
				For gintCurrentComponent = 0 To (gintGroupedComponents - 1)	'Execute all keywords in the group for the current group iteration
'					CRAFT_ReportEvent gstrReportedEventSheet,_ 
'						"Start Component", "Invoking Business component: " &_
'						garrComponentGroup(gintCurrentComponent), "Completed"
					
					CRAFT_ReportEvent gstrReportedEventSheet,_ 
						"Start : "& garrComponentGroup(gintCurrentComponent), "Invoking Business component: " &_
						garrComponentGroup(gintCurrentComponent), "Completed"
						
					'Check if the current keyword has already been invoked earlier, and update the hash table accordingly
					If gobjHashTable.Exists(garrComponentGroup(gintCurrentComponent)) Then
						gobjHashTable.Item(garrComponentGroup(gintCurrentComponent)) = _
							gobjHashTable.Item(garrComponentGroup(gintCurrentComponent)) + 1
					Else
						gobjHashTable.Add garrComponentGroup(gintCurrentComponent), 1
					End If
					
					
					'Update the current sub iteration number
					Environment.Value("CurrentSubIteration") = gobjHashTable._
								Item(garrComponentGroup(gintCurrentComponent))
										
					CRAFT_InvokeBusinessComponent garrComponentGroup(gintCurrentComponent)
					
'					CRAFT_ReportEvent gstrReportedEventSheet, "End Component",_
'											"Exiting Business component: " &_
'							garrComponentGroup(gintCurrentComponent), "Completed"
					
					'Changes made for Reporting on 24-June 2015, by Ram
					CRAFT_ReportEvent gstrReportedEventSheet, "End : "& garrComponentGroup(gintCurrentComponent),_
											"Exiting Business component: " &_
							garrComponentGroup(gintCurrentComponent), "Completed"
							
					If (Environment.Value("ExitIteration")) Then
						Exit Do
					End If
					
					If (Environment.Value("StopExecution")) Then
						CRAFT_ReportEvent gstrReportedEventSheet,_ 
							"CRAFT_Info", "Execution aborted by user", "Completed"
						CRAFT_CalculateExecTime()
						CRAFT_WrapUp gstrDatatableName
						ExitRun
					End If
				Next
			Next
			
			gintGroupedComponents = 0
		End If
		
		'Process next keyword
		gintCurrentFlowNumber = gintCurrentFlowNumber + 1
		garrCurrentFlowData = Split(DataTable.Value("Keyword_" &_
							gintCurrentFlowNumber,gstrBusinessFlowSheet),",")
		If UBound(garrCurrentFlowData) = -1 Then
			gstrCurrentKeyword = ""
		Else
			gstrCurrentKeyword = garrCurrentFlowData(0)
		End If
	Loop
	
	CRAFT_ReportEvent gstrReportedEventSheet, "End", "Iteration" &_
									gintCurrentIteration & " completed", "Completed"
	CRAFT_CalculateExecTime()
	
	'Move to the next iteration of test data
	gintCurrentIteration = gintCurrentIteration + 1
	gobjHashTable.RemoveAll()
Loop

Set gobjHashTable = Nothing

CRAFT_WrapUp gstrDatatableName


'*******************************************************************
'Sub QualifyLifeEvents()
'    On Error Resume Next
'    
'    
'    
	Set gobjPath = Browser("title:=.*Individual & Families.*").page("title:=.*Individual & Families.*")
    	Browser("title:=.*Individual & Families.*").sync
    	TINYWAIT
    
    	If gobjPath.WebElement("innertext:= Qualifying Life Events").Exist(2)Then
    
	    's	LifeEvents = "true::Checkbox-05/05/2016-false-true:Checkbox-05/05/2016-false-true~false~false~false"
	    sLifeEvents = "false~true:true::02/02/2015-false-false~false~false"
	    'CRAFT_GetData ("IndividualApp_Portal_Data","Applicant1_LifeEvents")'"false~false~false~true:08/01/2015:08/01/2015"
    
    		aLifeEvents = Split(sLifeEvents, "~")
    
    		For iLoop = LBound(aLifeEvents) To UBound(aLifeEvents) Step 1
			aItemSelect = Split(aLifeEvents(iLoop),":")
			Select Case (iLoop + 1)
				Case 1
		                	'Perform First Radio Box Selection
		                	SelectRadioButton "html id:=elgModification.heathCovChangeInHh.*",aItemSelect(0)
		                	TINYWAIT
		                
		                	'Based on input selection true handle the code as below
		                	For jLoop = LBound(aItemSelect) + 1 To UBound(aItemSelect) Step 1
		                	    	'Select the applicant input..
		                    		aApplicantData = Split(aItemSelect(jLoop), "-")
		                    
		                    		'Select for Checkbox, TextBox, RadioGroup
			                    	For kLoop = LBound(aApplicantData) To UBound(aApplicantData) Step 1
							If Trim(aApplicantData(kLoop)) <> "" Then
			                           		Select Case (kLoop + 1)
				                                'Selecting the Checkbox
									Case 1
										ClickCheckBox "name:=elgModification.elgMemberModifications\[" & UBound(aItemSelect) - jLoop & "\].lostOtherCoverage"
										TINYWAIT
				                                    
									Case 2
					                                   EnterData "html id:=elgModification.elgMemberModifications\[" & UBound(aItemSelect) - jLoop & "\].hltCoverageEndDate.*", aApplicantData(kLoop), "Enter Coverage EndDate"
					                                   TINYWAIT
					                                    
									Case 3                                
										SelectRadioButton "html id:=elgModification.elgMemberModifications" & UBound(aItemSelect) - jLoop & ".notPaidPremium.*", aApplicantData(kLoop)
					                                   TINYWAIT
					                                    
									Case 4
										SelectRadioButton "html id:=elgModification.elgMemberModifications" & UBound(aItemSelect) - jLoop & ".cancelCoverage.*", aApplicantData(kLoop)
										TINYWAIT
								End Select
			                        	End If
			                    Next
		                	Next
                
				Case 2
					'Perform Second Radio Box Selection
					SelectRadioButton "html id:=elgModification.dependentChangeInHh.*",aItemSelect(0)
					TINYWAIT
			                
					For jLoop = LBound(aItemSelect) + 1 To UBound(aItemSelect) Step 1
			                    'Select the applicant input..
						aApplicantData = Split(aItemSelect(jLoop), "-")
			                    
			                    'Select for Checkbox, TextBox, RadioGroup
			                    	For kLoop = LBound(aApplicantData) To UBound(aApplicantData) Step 1
			                        	If Trim(aApplicantData(kLoop)) <> "" Then
			                            	Select Case (kLoop + 1)
			                            	    'Selecting the Checkbox
			                                		Case 1
										SelectRadioButton "html id:=elgModification.marriageInHh.*", aApplicantData(kLoop)
										TINYWAIT
										
										ClickCheckBox "name:=elgModification.elgMemberModifications\[" & UBound(aItemSelect) - jLoop & "\].married"
										TINYWAIT
										
										EnterData "html id:=elgModification.elgMemberModifications" & UBound(aItemSelect) - jLoop & ".dateOfMarriage", aApplicantData(kLoop), "Date entered"
										TINYWAIT
			                                    
									Case 2
									    	SelectRadioButton "name:=elgModification.birthInHh", aApplicantData(kLoop)
									    
									Case 3    
									    	SelectRadioButton "name:=elgModification.fosterCareInHh", aApplicantData(kLoop)

			                                End Select
			                        End If
			                    Next
			                    
			                Next    
                
                
            Case 3
                'Perform Third Radio Box Selection
                SelectRadioButton "html id:=elgModification.immigrationChangeInHh.*",aItemSelect(0)
                TINYWAIT
                
            Case 4
                'Perform Fourth Radio Box Selection
                SelectRadioButton "html id:=elgModification.addressChangeInHh.*",aItemSelect(0)
                TINYWAIT
                
                'If strcomp(aLifeEvents(iLoop), "TRUE", 1) = 0 AND UBOUND(aItemSelect) > 0 Then
                If Instr(1, aLifeEvents(iLoop), "TRUE", 1) > 0 AND UBOUND(aItemSelect) > 0 Then
                    
                    iApplicants = 2' Cint(CRAFT_GetData ("IndividualApp_Portal_Data","MemberCount"))
    
                        'Loop to enter the details for all the Applicant 1, 2
                        For jLoop = 1 To iApplicants Step 1
                            If aItemSelect(jLoop) <> "" Then
                                ClickCheckBox "name:=elgModification.elgMemberModifications\["& (jLoop -1) &"\].memberMovedMA"
                                TINYWAIT
                        
                                If UBound(aItemSelect) <= iLoop Then
                                       'Enter the Data
                                         EnterData "html id:=elgModification.elgMemberModifications"& (jLoop - 1) &".moveToMaDt", aItemSelect(jLoop), "Qualifying Life Events --> HouseHold Status"
                                          TINYWAIT
                                End If
                            End If
                        Next
                End If
            Case Else
                Msgbox "Code Handled only for Radio Box Selection till - 4 in Function : --> "& QualifyLifeEvents
        End Select
    Next
    End If




'
'End Sub
