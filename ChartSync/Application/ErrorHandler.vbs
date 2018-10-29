'**********************************************************************************************
'**********************************************************************************************
' errorHandlers.vbs
' Functions - See individual Function Headers for
'             a detail description of function functionality
'       appTimeout			'click the "Continue" button on the popup window time out message if exists.
'       cannotDisplayPage
'       requestedURL
'       pageExpired
'       errorOnPage
'       
'**********************************************************************************************
'**********************************************************************************************
Option Explicit

'<@class_summary>
'**********************************************************************************************
' <@purpose>
'   This Class is used to interact with Application Errors thrown such as Page Cannot Be Displayed
'   or requested URL not found or Page Expired Errors.
'   Execution of this Class File will create a errorHandlers Object automatically
'   Object Name equals "eh"
' </@purpose>
'
' <@author>
'		Mike Millgate
' </@author>
'
' <@creation_date>
'   11-07-2007
' </@creation_date>
'
'**********************************************************************************************
'</@class_summary>
Class errorHandlers

	'<@comments>
	'**********************************************************************************************
	'	<@name>cannotDisplayPage</@name>
	'
	' <@purpose>To check for "the page cannot be displayed" error in a browser window
	'             If found clicks the browser back button to return to a known good state
	' </@purpose>
	'
	' <@parameters>None</@parameters>
	'
	' <@return>Boolean Value (True/False)
	'            True -  If found
	'            False - If not found or other function errors
	' </@return>
	'
	'	<@assumptions>Environment Variable "WINDOW_ID" has been initialized/created</@assumptions>
	'
	'	<@example_usage>eh.cannotDisplayPage()</@example_usage>
	'
	' <@author>Mike Millgate</@author>
	'
	' <@creation_date>11-07-2007</@creation_date>
	'
	' <@mod_block>
	'	   11-28-2007 - MM - Added Browser(oAppBrowser).Back to put the application
	'                      back to a known good state
	'    03-18-2008 - MM - Added Logic to Clear Object Variables to free memory
	'	</@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Function cannotDisplayPage() ' <@as> Boolean
	
	  Services.StartTransaction "cannotDisplayPage" ' Timer Begin   
	  Reporter.ReportEvent micDone, "cannotDisplayPage Function", "Function begin"
	   
	  ' Variable Declaration / Initialization      
	  Dim oAppBrowser, oPCDError
		
		' Description Object Declarations/Initializations
	  Set oAppBrowser = Description.Create()
	  oAppBrowser("MicClass").Value = "Browser"
	  oAppBrowser("Hwnd").Value = Environment.Value("WINDOW_ID")
	   
	  Set oPCDError = Description.Create()
	  oPCDError("MicClass").Value = "WebElement"
	  oPCDError("innertext").Value = "The page cannot be displayed"
	  oPCDError("html id").Value = "errorText"
	
	  ' Verification of the Object
	  If Browser(oAppBrowser).WebElement(oPCDError).Exist(1) Then
	  	cannotDisplayPage = True ' Return Value
	    ' Click the Browser Back Button to put application back to a known good state
	    Browser(oAppBrowser).Back
	  Else ' Object Not Found
	    cannotDisplayPage = False ' Return Value
	  End If
	   
	  ' Clear object variables
	  Set oAppBrowser = Nothing
	  Set oPCDError = Nothing
	
	  Reporter.ReportEvent micDone, "cannotDisplayPage Function", "Function End"   
	  Services.EndTransaction "cannotDisplayPage" ' Timer End
	
	End Function

	'<@comments>	
	'**********************************************************************************************
	' <@name>requestedURL</@name>
	'
	' <@purpose>To check for "the requested URL could not be retrieved" error in a browser window
	'             If found clicks the browser back button to return to a known good state
	' </@purpose>
	'
	' <@parameters>None</@parameters>
	'
	' <@return>Boolean Value (True/False)
	'            True -  If found
	'            False - If not found or other function errors
	' </@return>
	'
	' <@assumptions>Environment Variable "WINDOW_ID" has been initialized/created</@assumptions>
	'
	' <@example_usage>eh.requestedURL()</@example_usage>
	'
	' <@author>Mike Millgate</@author>
	'
	'	<@creation_date>12-05-2007</@creation_date>
	'
	' <@mod_block>
	'   03-18-2008 - MM - Added Logic to Clear Object Variables to free memory  
	' </@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Function requestedURL() ' <@as> Boolean
	
	  Services.StartTransaction "requestedURL" ' Timer Begin   
	  Reporter.ReportEvent micDone, "requestedURL Function", "Function begin"
	   
	  ' Variable Declaration / Initialization      
	  Dim oAppBrowser, oError
	
	  ' Description Object Declarations/Initializations
	  Set oAppBrowser = Description.Create()
	  oAppBrowser("MicClass").Value = "Browser"
	  oAppBrowser("Hwnd").Value = Environment.Value("WINDOW_ID")
	  
	  Set oError = Description.Create()
	  oError("MicClass").Value = "WebElement"
	  oError("innertext").Value = "The requested URL .* could not be retrieved"
	  oError("html tag").Value = "H2"
	
	  ' Verification of the Object
	  If Browser(oAppBrowser).WebElement(oError).Exist(1) Then
	    requestedURL = True ' Return Value
	    ' Click the Browser Back Button to put application back to a known good state
	    Browser(oAppBrowser).Back
	  Else ' Object Not Found
	    requestedURL = False ' Return Value
	  End If
	   
	  ' Clear object variables
	  Set oAppBrowser = Nothing
	  Set oError = Nothing
	
	  Reporter.ReportEvent micDone, "requestedURL Function", "Function End"   
	  Services.EndTransaction "requestedURL" ' Timer End
	
	End Function
	
	'<@comments>
	'**********************************************************************************************
	' <@name>pageExpired</@name>
	'
	' <@purpose>To check for "the Page Has Expired" error in a browser window
	'           If found clicks the browser back button to return to a known good state, if possible
	' </@purpose>
	'
	' <@parameters>None</@parameters>
	'
	' <@return>Boolean Value (True/False)
	'            True -  If found
	'            False - If not found or other function errors
	' </@return>
	'
	' <@assumptions>Environment Variable "WINDOW_ID" has been initialized/created</@assumptions>
	'
	' <@example_usage>eh.pageExpired()</@example_usage>
	'
	' <@author>Mike Millgate</@author>
	'
	' <@creation_date>04-04-2008</@creation_date>
	'
	' <@mod_block></@mod_block>
	'
	'**********************************************************************************************
	'</@comments>	
	Public Function pageExpired() ' <@as> Boolean
	
	  Services.StartTransaction "pageExpired" ' Timer Begin   
	  Reporter.ReportEvent micDone, "pageExpired Function", "Function begin"
	   
	  ' Variable Declaration / Initialization      
	  Dim oAppBrowser, oError
	
	  ' Description Object Declarations/Initializations
	  Set oAppBrowser = Description.Create()
	  oAppBrowser("MicClass").Value = "Browser"
	  oAppBrowser("Hwnd").Value = Environment.Value("WINDOW_ID")
	   
	  Set oError = Description.Create()
	  oError("MicClass").Value = "WebElement"
	  oError("innertext").Value = "Warning: Page has Expired.*"
	  oError("html tag").Value = "FONT"
	
	   ' Verification of the Object
	  If Browser(oAppBrowser).WebElement(oError).Exist(1) Then
	  	pageExpired = True ' Return Value
	  	' Click the Browser Back Button to put application back to a known good state, if possible
	  	Browser(oAppBrowser).Back
	  Else ' Object Not Found
	  	pageExpired = False ' Return Value
	  End If
	   
	  ' Clear object variables
	  Set oAppBrowser = Nothing
	  Set oError = Nothing
	
	  Reporter.ReportEvent micDone, "pageExpired Function", "Function End"   
	  Services.EndTransaction "pageExpired" ' Timer End
	
	End Function
	
	'<@comments>
	'**********************************************************************************************
	' <@name>errorOnPage</@name>
	'
	' <@purpose>
	'    To check for "Error On Page" script error dialog window is thrown
	'      If found clicks the ok button of the dialog window
	'
	'    To Note: This function checks for the dialog box thrown, for this to happen
	'             the Internet Explorer setting to show dialog when script errors
	'             found must be turned on
	'             Internet Options Advanced - Display a notification about every script error
	' </@purpose>
	'
	' <@parameters>None</@parameters>
	'
	' <@return>None</@return>
	'
	' <@assumptions>Environment Variable "WINDOW_ID" has been initialized/created</@assumptions>
	'
	' <@example_usage>eh.errorOnPage()</@example_usage>
	'
	' <@author>Mike Millgate</@author>
	'
	' <@creation_date>05-15-2008</@creation_date>
	'
	' <@mod_block>
	' </@mod_block>
	'
	'**********************************************************************************************
	'</@comments>
	Public Sub errorOnPage() ' <@as> Nothing
	
	   Services.StartTransaction "errorOnPage" ' Timer Begin   
	   Reporter.ReportEvent micDone, "errorOnPage Function", "Function begin"
	   
	   ' Variable Declaration / Initialization      
	   Dim oAppBrowser, oDialog, oText, oButton, oErrorText, sErrorText
	
	   ' Description Object Declarations/Initializations
	   Set oAppBrowser = Description.Create()
	   oAppBrowser("MicClass").Value = "Browser"
	   oAppBrowser("Hwnd").Value = Environment.Value("WINDOW_ID")
	   
	   ' Dialog Object
	   Set oDialog = Description.Create()
	   oDialog("micClass").Value = "Window"
	   oDialog("text").Value = "Internet Explorer.*"
	   oDialog("is owned window").Value = True
	   oDialog("is child window").Value = False
	   oDialog("nativeclass").Value = "Internet Explorer_TridentDlgFrame"
	   
	   ' Text Object
	   Set oText = Description.Create()
	   oText("micclass").Value = "WebElement"
	   oText("html id").Value = "tdMsg"
	   oText("innertext").Value = "Problems with this Web page might prevent it from being " _
	                              & "displayed properly or functioning properly. In the future, " _
	                              & "you can display this message by double-clicking the warning " _
	                              & "icon displayed in the status bar. "
	   
	   ' Button Object
	   Set oButton = Description.Create()
	   oButton("micclass").Value = "WebButton"
	   oButton("html id").Value = "btnOK"
	   
	   ' Error Text Object
	   Set oErrorText = Description.Create()
	   oErrorText("micclass").Value = "WebTable"
	   oErrorText("html id").Value = "tbl2"
	   
	   ' Verification of the Object
	   If Browser(oAppBrowser).Window(oDialog).WebElement(oText).Exist(3) Then
	   				 ' Get the Error Text
	   				 sErrorText = Browser(oAppBrowser).Window(oDialog).WebTable(oErrorText).GetROProperty("innertext")
	   				 Reporter.ReportEvent micFail, "Error On Page", "Error On Page Script Error Found!" _
	   				                               & vbNewLine & "Error Message" & vbNewLine & sErrorText
	           ' Click the OK Dialog Button
	           Browser(oAppBrowser).Window(oDialog).WebButton(oButton).Click
	   End If
	   
	   ' Clear object variables
	   Set oAppBrowser = Nothing
	   Set oDialog = Nothing
	   Set oText = Nothing
	   Set oButton = Nothing
	   SEt oErrorText = Nothing
	
	   Reporter.ReportEvent micDone, "errorOnPage Function", "Function End"   
	   Services.EndTransaction "errorOnPage" ' Timer End	
	End Sub

	Public Sub appTimeout()
	'********************************************************************************************
	'Purpose: Click the Continue button on the popup window app. time out message if exists.
	'Parameters: None
	'Requires: Environment("BROWSER_OBJ") MUST EXISTS
	'Returns: None
	'Usage: eh.appTimeout()
	'Created by: Hung Nguyen 12/2/10
	'Modified:
	'********************************************************************************************
	   Services.StartTransaction "appTimeout" ' Timer Begin       
	   Dim oAppBrowser, oDialog, sTimeoutMesg
	   sTimeoutMesg="Your session is about to time out; Please press continue to renew the session"
	
	   '  Browser obj
	   Set oAppBrowser = Browser("title:=Ingenix InSite")   'Environment("BROWSER_OBJ"))
	
		If oAppBrowser.WebElement("html tag:=P","innertext:=" & sTimeoutMesg).Exist(3) Then
	
			' Important: making sure the obj is visible within 10 secs prior to clicking the Continue button
			If oAppBrowser.WebElement("html tag:=P","innertext:=" & sTimeoutMesg).WaitProperty("visible",True,1000)  Then 	
				If oAppBrowser.WebButton("html tag:=INPUT","type:=button","name:=Continue","index:=0").Exist(2) Then
					oAppBrowser.WebButton("html tag:=INPUT","type:=button","name:=Continue","index:=0").Click
					Reporter.ReportEvent micDone,"Timeout Message","The Continue button was clicked."
				Else
					Reporter.ReportEvent micWarning,"Timeout Message","The Continue button does not exist."
				End If 
			Else
				Reporter.ReportEvent micWarning,"Timeout Message","The popup window time out messsage is not visible. Please wait..."
			End If			
		Else
			Reporter.ReportEvent micWarning,"Timeout Message","The time out messsage '" &sTimeoutMesg &"' does not exist. Nothing to do."
		End If
		
	   ' Clear object 
	   Set oAppBrowser = Nothing   
	   Services.EndTransaction "appTimeout" ' Timer End
	End Sub

End Class

'**********************************************************************************************
'*                            Class Instantiation                                         
'**********************************************************************************************
dim eh

set eh = new errorHandlers