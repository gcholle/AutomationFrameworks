'**********************************************************************************************
'**********************************************************************************************
' logout.vbs
' Functions - See individual Function Headers for
'             a detail description of function functionality
'       logoutClose
'       logout
'       
'**********************************************************************************************

'**********************************************************************************************
'**********************************************************************************************
Option Explicit

Class logoutFunctions
	Function logoutClose()
	'**********************************************************************************************
	' Purpose: Click the Sign Out link to logout and close the IE Browser 
	' Parameters:  None
	' Returns: True/False
	' Assumptions:  Environment Variable BROWSER_OBJ has been initialized/created
	' Example Usage:  logoutClose()
	' Created by: Sujatha Kandikonda on 03/24/2011
	' Modified: Hung Nguyen 7/19/11 - Added to trap error, do loop to wait for the Confirm popup window
	'                                 prior to click the Confirm button.
	'			Govardhan 08/09/2012 - Updated the Link Signout with additional "innertext" property
	'           Hung Nguyen 09/10/12 - Updated due to the WebElement obj property changed for the Sign Out Confirming Dialog window
	 '                                 Updated the Confirm button obj w/index value
	 '          Hung Nguyen 3/29/13 - Updated to return True if the Sign In button appears from the login page
	'**********************************************************************************************
	   Services.StartTransaction "logoutClose" ' Timer Begin
	   Dim oAppBrowser, oLink,iTimeout,cnt,iClose
	   iTimeout=15	'secs
	   logoutClose=False	'init value
	
		Set oAppBrowser =Browser(Environment("BROWSER_OBJ"))			'Browser obj 

	   'Set oLink=oAppBrowser.Link("html tag:=A","text:=Sign Out")	'the Sign Out link obj.
	   Set oLink=oAppBrowser.Link("html tag:=A","innertext:=Sign Out","text:=Sign Out")	'the Sign Out link obj.
		
		' Logout 
		If oLink.Exist(2) Then 'exists
			oLink.Click ' Click the Sign Out Link
	
			'loop 15 secs waiting for the confirm popup window to click the Confirm button
			cnt=0
			Do while cnt <= iTimeout
				Wait(1)
				cnt=cnt+1

				'the dialog window WebElement
				If oAppBrowser.WebElement("html id:=confirmLogoutModalPanelContentDiv","innertext:=Sign Out.*").Exist(1) Then
					reporter.ReportEvent micInfo,"The Sign Out Confirming Dialog window exists",""
					If oAppBrowser.WebButton("html tag:=INPUT","name:=Confirm","type:=submit","index:=1").Exist(1) Then		'the Confirm button w/index
						oAppBrowser.WebButton("html tag:=INPUT","name:=Confirm","type:=submit","index:=1").Click	'click the button
						oAppBrowser.Sync
						Exit Do						
					End If
				End If
			Loop
	
			'report
			If oAppBrowser.WebButton("html id:=loginForm:cmdBtn").Exist(5) Then
				oAppBrowser.Close		'now close the browser window
				Reporter.ReportEvent micDone, "logoutClose", "Sign out was successful."
				logoutClose=True
			Else	
				Reporter.ReportEvent micWarning, "logoutClose", "Sign out was not successful."		'do not fail - warning only
			End If
		Else
			Reporter.ReportEvent micInfo, "logoutClose", "The Sign Out Link was not found. Closing the Browser window..."
			logoutClose=True
			If oAppBrowser.Exist(2) Then oAppBrowser.Close
		End If
		Set oLink=Nothing 
		Set oAppBrowser=Nothing
		Services.EndTransaction "logoutClose" ' Timer End
	End Function

	Function logout()
	'**********************************************************************************************
	' Purpose: Click the Sign Out link to logout and verify the Sign In page returns 
	' Parameters:  None
	' Returns: True/False
	' Assumptions:  Environment Variable BROWSER_OBJ has been initialized/created
	' Example Usage:  lout.logout()
	' Created by: Sujatha Kandikonda on 03/24/2011
	' Modified: Hung Nguyen 7/19/11 - Added to trap error, do loop to wait for the Confirm popup window
	'           prior to click the Confirm button.
	'			Govardhan 08/09/2012 - Updated the Lick Signout with additional "innertext" property
	'           Hung Nguyen 09/10/12 - Updated due to the WebElement obj property changed for the Sign Out Confirming Dialog window
	'                                  Updated the Confirm button obj w/index value		
	'**********************************************************************************************
	   Services.StartTransaction "logout" ' Timer Begin
	   Dim oAppBrowser, oLink,iTimeout,cnt,iClose
	   iTimeout=30	'secs
	   logout=False	'init value
	
		Set oAppBrowser =Browser(Environment("BROWSER_OBJ"))			'Browser obj 

	   'Set oLink=oAppBrowser.Link("html tag:=A","text:=Sign Out")	'the Sign Out link obj.
	   Set oLink=oAppBrowser.Link("html tag:=A","innertext:=Sign Out","text:=Sign Out")	'the Sign Out link obj.
		
		' Logout 
		If oLink.Exist(2) Then 'exists
			oLink.Click ' Click the Sign Out Link
	
			'loop 30 secs waiting for the confirm popup window to click the Confirm button
			cnt=0
			iClose=0
			Do while cnt <= iTimeout
				Wait(1)
				cnt=cnt+1

				'the dialog window WebElement
				If oAppBrowser.WebElement("html id:=confirmLogoutModalPanelContentDiv","innertext:=Sign Out.*").Exist(1) Then
					reporter.ReportEvent micInfo,"The Sign Out Confirming Dialog window exists",""
					If oAppBrowser.WebButton("html tag:=INPUT","name:=Confirm","type:=submit","index:=1").Exist(1) Then		'the Confirm button w/index
						oAppBrowser.WebButton("html tag:=INPUT","name:=Confirm","type:=submit","index:=1").Click	'click the button
						oAppBrowser.Sync
						iClose=1
						Exit Do						
					End If
				End If
			Loop
	
			'report
			If iClose=1 Then
				Reporter.ReportEvent micDone, "logout", "Sign out was successful."
				logout=True
			Else	
				Reporter.ReportEvent micWarning, "logout", "Sign out was not successful."		'do not fail - warning only
			End If
		Else
			Reporter.ReportEvent micFail, "logout", "The Sign Out Link was not found. Unable to sign out."
		End If
		Set oLink=Nothing 
		Set oAppBrowser=Nothing
		Services.EndTransaction "logout" ' Timer End
	End Function
End Class
'**********************************************************************************************
'*                            Class Instantiation                                         
'**********************************************************************************************
dim lout

set lout = New logoutFunctions