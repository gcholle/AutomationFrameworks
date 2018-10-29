'**********************************************************************************************
'**********************************************************************************************
' login.vbs
' Functions - See individual Function Headers for
'             a detail description of function functionality
'       login
'       loginUser
'       ResetPassword
'		OpenBrowserInstance
'**********************************************************************************************
'**********************************************************************************************
Option Explicit

Class Login
	Public Function login() 
	'*******************************************************************************************************
	' Purpose: ChartSync - Admin login via Environment("ADMIN") and Environment("CHARTSYNC_PWD") from the initialize file
	' Parameters: None
	' Requires: ChartSync_Initialize.vbs must be executed for the global environments to be available
	' Calls: objFunctions.vbs
	' Returns: True/False
	' Usage: lin.login()
	' Created by: Hung Nguyen 10/25/10
	' Modified: Hung Nguyen 7/19/11 - Updated to work w/ChartSync.
	'			Govardhan Choletti 2/16/2012 - Updated the Browser object - oAppBrowser correctly
	'*******************************************************************************************************
		Services.StartTransaction "login" ' Timer Begin
		Dim sUser,sPwd, oAppBrowser,cnt,iLogin,iTimeout
		iTimeout=120 '120 secs

		login=False	'init value
	
		'verify global environment values
		Err.Clear
		On error resume next 
		sUser=Environment("ADMIN")
		sPwd=Environment("CHARTSYNC_PWD")
		Set oAppBrowser = Browser(Environment("BROWSER_OBJ"))			' Browser obj
		If Err.Number <> 0 Then
			reporter.ReportEvent micFail,"login","Error number = " &Err.Number _
			                                                          &vbnewline &"Error description = " &Err.Description
			Exit Function
		End If
		On error goto 0
		
		'If  "Go to Sign In" Link is shown on page beacuse of delay in Logging
		If of.linkClicker("Go to Sign In") = True Then
			Reporter.ReportEvent micPass,"Link : Go to Sign In","Successfully Clicked on Link 'Go to Sign In' Page as Expected"
		Else
			Reporter.ReportEvent micInfo,"Link : Go to Sign In","Link 'Go to Sign In' Page is Not displayed on Screen"
		End If

		'start login
		If oAppBrowser.Exist Then	
			If of.webEditEnter("loginForm:usrNameValue",sUser) Then 	'enter user id
				Reporter.ReportEvent micDone,"Enter Value","Function call to enter value '" &sUser &"' into field was successful."
				If of.webEditEnter("loginForm:passwdValue",sPwd) Then 'enter password
					Reporter.ReportEvent micDone,"Enter Value","Function call to enter value '" &sPwd &"' into field was successful."
					If of.imageClicker("sign_in.png") OR of.webButtonClicker("loginForm:cmdBtn") = True Then	'click the Sign In image
						Reporter.ReportEvent micDone,"login","Function call to click 'Sign In' was successful."
						
						'loop 2 min. waiting
						cnt=0
						iLogin=0
						Do While cnt <= iTimeout
							Wait(1)
							cnt=cnt+1
							
							''invalid username or password
							If oAppBrowser.WebElement("innertext:=Invalid Username or Password","html tag:=SPAN").Exist(1) Then 
								Reporter.ReportEvent micWarning,"login","Invalid user name or password."
								Exit Do 
							ElseIf oAppBrowser.Link("innertext:=Sign Out","html tag:=A").Exist(1) Then		'Sign Out link exists
								iLogin=1
								Exit Do
							End If 
						Loop
												
						'report
						If iLogin=1 Then
							Reporter.ReportEvent micDone,"login","User '" &sUser &"' was logged in successfully."
							login=True
						Else
							Reporter.ReportEvent micWarning,"login","User '" &sUser &"' was not logged in successfully."
						End If
					Else
						Reporter.ReportEvent micWarning,"login","Function call to click 'Sign In' was not successful."
					End If  			
				Else
					Reporter.ReportEvent micWarning,"login","Function call to enter value '" &sPwd &"' into field was not successful."
				End If 
			Else
				Reporter.ReportEvent micWarning,"login","Function call to enter value '" &sUser &"' into field was not successful."
			End If
		Else
			Reporter.ReportEvent micWarning,"login","The login browser page does not exist. Unable to login."
		End If

		' Clear Object
		Set oAppBrowser = Nothing
		Services.EndTransaction "login" ' Timer End  	
	End Function
	
	Public Function loginUser(sUserId,sPassword) 
	'*******************************************************************************************************
	' Purpose: ChartSync - User login 
	' Parameters: sUserId = string - user id
	'             sPassword =  string - password
	' Requires: ChartSync_Initialize.vbs must be executed for the global environments to be available
	' Calls: objFunctions.vbs
	' Returns: True/False
	' Usage: lin.loginUser("me","mypassword")
	' Created by: Hung Nguyen 10/25/10
	' Modified: Hung Nguyen 7/19/11 - Added to check for invalid user name or password message and Environment BROWSER_OBJ existence
	'			Govardhan Choletti 1/11/2012 -  Added code to Handle Dialog Security Alert Box in STAGE Environment
	'			Govardhan Choletti 2/3/2012 - Added code to handle Reset Password for any user Role
	'			Govardhan Choletti 2/16/2012 - Updated the Browser object - oAppBrowser correctly
	'           Hung Nguyen 10/26/12 - UPdated use obj html id value (not calling function from objFunction file) for entering userid and password.
	'*******************************************************************************************************
		Services.StartTransaction "loginUser" ' Timer Begin
		Dim oAppBrowser,cnt,iLogin,iTimeout,bPwdReset
		iTimeout=120 '120 secs
		loginUser=False	'init value

		'verify parameters
		If isempty(sUserId) or sUserID="" or isempty(sPassword) or sPassword="" Then
			reporter.ReportEvent micFail,"loginUser","Invalid parameter passed. Please try again."
			Exit Function
		End If

		'verify env. var.
		On error resume next 
			Set oAppBrowser = Browser(Environment("BROWSER_OBJ"))			' Browser obj
			oAppBrowser.Sync
			
			If Err.Number <> 0 Then
				reporter.ReportEvent micFail,"loginUser","Environment BROWSER_OBJ is not available. Error number = " &Err.Number _
				                                                          &vbnewline &"Error description = " &Err.Description
				Exit Function
			End If
		On error goto 0
		
		If oAppBrowser.Exist(10) Then	
			If oAppBrowser.WebEdit("html id:=loginForm:usrNameValue").Exist(5) Then 	'enter user id
				oAppBrowser.WebEdit("html id:=loginForm:usrNameValue").Set sUserid
				If oAppBrowser.WebEdit("html id:=loginForm:passwdValue").Exist(5) Then 	'enter password
					oAppBrowser.WebEdit("html id:=loginForm:passwdValue").set sPassword
					If oAppBrowser.WebButton("html id:=loginForm:cmdBtn").Exist(5) Then
						oAppBrowser.WebButton("html id:=loginForm:cmdBtn").Click		'click the Sign In button				
						
						' Added code to Handle Dialog Security Alert Box in STAGE Environment
						Dim oFso, sIEVersion, iIEVersion, iCnt
						If Instr(1, Environment("TEST_BROWSER_URL"), "stage", 1) > 0 Then
							Set oFso = CreateObject("Scripting.FileSystemObject") 
							sIEVersion = oFso.GetFileVersion("C:\Program Files\Internet Explorer\iexplore.exe") 
							iIEVersion=CInt(Left(sIEVersion,1))
							iCnt = 0
							If iIEVersion < 7 Then
								If oAppBrowser.Dialog("regexpwndtitle:=Security Alert").Exist(2) Then
									Do
										oAppBrowser.Dialog("regexpwndtitle:=Security Alert").WinButton("text:=&Yes").Click
										iCnt = iCnt + 1
									Loop Until (oAppBrowser.Dialog("regexpwndtitle:=Security Alert").Exist(1) = False OR iCnt = 60)
								End If
							End If	
							Set oFso=Nothing
						End If
						
						'loop 2 min. waiting
						cnt=0
						iLogin=0
						Do While cnt <= iTimeout
							Wait(1)
							cnt=cnt+1
							
							'invalid username or password
							If oAppBrowser.WebElement("innertext:=Invalid Username or Password","html tag:=SPAN").Exist(1) Then 
								Reporter.ReportEvent micFail,"loginUser","Invalid user name or password."
								Exit Do 
							ElseIf oAppBrowser.Link("innertext:=Sign Out","html tag:=A").Exist(1) Then		'Sign Out link exists
								iLogin=1
								Exit Do
							End If 
						Loop
						
						' Check for the Password Reset Criteria is displayed on screen
						bPwdReset = lin.ResetPassword (sUserId,sPassword)
						
						'report
						If iLogin=1 And bPwdReset Then
							Reporter.ReportEvent micDone,"loginUser","User '" &sUserId &"' was logged in successfully."
							loginUser=True
						Else
							Reporter.ReportEvent micWarning,"loginUser","User '" &sUserId &"' was not logged in successfully."
						End If
					Else
						'Reporter.ReportEvent micWarning,"loginUser","Function call to click 'Sign In' was not successful."
						Reporter.ReportEvent micWarning,"loginUser","The Sign In button was not clicked successfully."
					End If  			
				Else
					Reporter.ReportEvent micWarning,"loginUser","Password '" &sPassword &"' was not entered into field successfully."
					'Reporter.ReportEvent micWarning,"loginUser","Function call to enter value '" &sPassword &"' into field was not successful."
				End If 
			Else
				Reporter.ReportEvent micWarning,"loginUser","Userid '" &sUserId &"' was not entered into field successfully."
				'Reporter.ReportEvent micWarning,"loginUser","Function call to enter value '" &sUserId &"' into field was not successful."
			End If
		Else
			Reporter.ReportEvent micWarning,"loginUser","The login browser page does not exist. Unable to login."
		End If

		' Clear Object
		Set oAppBrowser = Nothing
		Services.EndTransaction "loginUser" ' Timer End  	
	End Function
	
	Public Function ResetPassword(sUserRoleID, sPassword) 
	'*******************************************************************************************************
	' Purpose: ChartSync - User login 
	' Parameters: sUserRoleID = string - user id
	'             sPassword =  string - password
	' Requires: ChartSync_Initialize.vbs must be executed for the global environments to be available
	' Calls: login.vbs, logout.vbs
	' Returns: True/False
	' Usage: lin.ResetPassword("test123", "password")
	' Created by: Govardhan Choletti 2/2/2012
	' Modified:   Govardhan Choletti - No Comments
	'			  Appended DB Table name with 'ravas.'
	'*******************************************************************************************************
	Services.StartTransaction "ResetPassword"
	Reporter.ReportEvent micDone, "ResetPassword", "Function Begin"

	Dim sQuery, sTempPassword
	Dim oRS
	sTempPassword = Environment("RESET_PWD") '"Password-123"
	ResetPassword = False
	'sUserRoleID = "test123"
	'sPassword = "password"
' If the Message 'Password Expired' is shown on screen, Then - Enter New Password and ReEnter Password
	If of.webElementFinder("passwdForm:expiredMsg") = True Then
		Reporter.ReportEvent micInfo,"STEP - Password Expired ","Message - 'Password expired, please change your password.' is Shown on Screen"

		' Verify 'New Password'  WebElement
		If of.webElementFinder("passwdForm:newPwd") = True Then
			Reporter.ReportEvent micInfo,"STEP - New Password*  ","Message - 'New Password*' field is Shown on Screen"

			'Enter any Valid Password in the New Password'  Field
			If of.webEditEnter("passwdForm:newPdValue", sTempPassword) = True Then
				Reporter.ReportEvent micInfo,"STEP - New Password*  Entry ","New Password - '"& sTempPassword &"' entered successfully"
			Else
				Reporter.ReportEvent micFail,"STEP - New Password*  Entry ","Unable to enter New Password - '"& sTempPassword &"'"
			End If

			' Verify 'Confirm New Password'  WebElement
			If of.webElementFinder("passwdForm:confNewPwd") = True Then
				Reporter.ReportEvent micInfo,"STEP - Confirm New Password*","Message - 'Confirm New Password*' field is Shown on Screen"

				'Enter any Valid Password in the 'Confirm New Password' Field
				If of.webEditEnter("passwdForm:confPdValue", sTempPassword) = True Then
					Reporter.ReportEvent micInfo,"STEP - Confirm New Password*  Entry ","Confirm New Password - '"& sTempPassword &"' entered successfully"
				Else
					Reporter.ReportEvent micFail,"STEP - Confirm New Password*  Entry ","Unable to enter Confirm New Password - '"& sTempPassword &"'"
				End If

				' Click on 'Change Password' Button
				If of.webButtonClicker("passwdForm:cmdBtn2") = True Then
					Reporter.ReportEvent micInfo,"STEP - Click - Button Change Password'","Successfully clicked on button 'Change Password'"

					' Validate the Message 
					Wait(2)
					If of.webElementFinder("passwdForm:changePwdSuccessMessage") = True Then
						Reporter.ReportEvent micInfo,"STEP - Password Changed ","Message - ' Password has been updated successfully. Please sign out and sign in once again with the new password.' is Shown on Screen"
					Else
						Reporter.ReportEvent micFail,"STEP - Password Changed ","No Message - Displayed on Screen --> Continuing with DB Reset"
					End If

					' Sign Out from the Application
                    lout.logout()

					' Validate the Password Expired Field in Data Base
					sQuery = "SELECT passwdexpired FROM ravas.app_user WHERE logonid='"& sUserRoleID &"'"
					Set oRS = db.executeDBQuery(sQuery, Environment.Value("DB"))
					If LCase(typeName(oRS)) <> "recordset" Then
						Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
					ElseIf oRS.bof And oRS.eof Then
						Reporter.ReportEvent micFail, "invalid recordset", "The returned recordset contains no records."
					ElseIf CStr(oRS.fields(0).Value) = "1" Then
						Reporter.ReportEvent micInfo,"STEP - Query, Check Password Expiry","Password Expired for the User -'"& sUserRoleID &"' as Expected"
					
						' Updating/Resetting the User  password in Data Base with 'password'
						sQuery = "UPDATE ravas.app_user "_
								&"SET logonpasswd='x3s507Zz6ylStlAcW7DCrhPShYTfc5QkCjZJ2Ju4b236f3MEMe6wjMBuTwdhEa0y', account_lock='0', passwdexpired='0', passwdretries='0' "_
								&"WHERE logonid='"& sUserRoleID &"'"
						Set oRS = db.executeDBQuery(sQuery, Environment.Value("DB"))
						If LCase(typeName(oRS)) <> "recordset" Then
							Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
						Else
							Reporter.ReportEvent micInfo,"STEP - Run Reset Query ","Successfully ran the Query - '"& sQuery &"' to reset the Password"
	
							' Commit the Query 
							sQuery = "COMMIT"
							Set oRS = db.executeDBQuery(sQuery, Environment.Value("DB"))
							If LCase(typeName(oRS)) <> "recordset" Then
								Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
							Else
								Reporter.ReportEvent micInfo,"STEP - Commit Reset Query ","Successfully comitted Query - 'COMMIT' to reset the Password"

								' Validate the Password resetted successfully or Not
								sQuery = "SELECT passwdexpired FROM ravas.app_user WHERE logonid='"& sUserRoleID &"'"
								Set oRS = db.executeDBQuery(sQuery, Environment.Value("DB"))
								If LCase(typeName(oRS)) <> "recordset" Then
									Reporter.ReportEvent micFail, "invalid recordset", "The database connection did not open or invalid parameters were passed."
								ElseIf oRS.bof And oRS.eof Then
									Reporter.ReportEvent micFail, "invalid recordset", "The returned recordset contains no records."
								ElseIf CStr(oRS.fields(0).Value) = "0" Then
									Reporter.ReportEvent micInfo,"STEP - Query - Validate Password ","Password Resetted to default password 'password' successfully as Expected"
									ResetPassword = True
								Else
									Reporter.ReportEvent micFail,"STEP - Query - Validate Password ","Unable to Reset the Password to default password, Not as Expected"
								End If
							End If
						End If
					End If

				' Sign in to the Application Once again with the User Supplied Login credentials
				 lin.loginUser sUserRoleID, sPassword
				Else
					Reporter.ReportEvent micFail,"STEP - Click - Button 'Change Password'","Unable to click on button 'Change Password'"
				End If
			Else
				Reporter.ReportEvent micFail,"STEP - Confirm New Password*","Message - 'Confirm New Password*' field is Not Shown on Screen"
			End If
		Else
			Reporter.ReportEvent micFail,"STEP - New Password*  ","Message - 'New Password*' field is Not Shown on Screen"
		End If
	Else
		Reporter.ReportEvent micInfo,"STEP - Password Expired ","No Message - Displayed on Screen --> Continuing with Script Execution"
		ResetPassword = True
	End If

' Release the memory allocated to the variable
	Set oRS = Nothing
		
' Close the transaction 
	Services.EndTransaction "ResetPassword"
	End Function
	
	Public Function OpenBrowserInstance(sBrwInstance) 
	'*******************************************************************************************************
	' Purpose: Create a New Browser Instance of URL Passed and Store Browser Object in the Environmental Variable
	' Parameters: sBrwInstance = "0" or "1" or "2"
	' Requires: ChartSync_Initialize.vbs must be executed for the global environments to be available
	' Calls: Nothing
	' Returns: True/False
	' Usage: lin.OpenBrowserInstance()
	' Created by: Govardhan Choletti 02/14/11
	' Modified: 
	'*******************************************************************************************************
	' Variable Declaration / Initialization
	Dim sUrl, oAppBrowserInst, windowID,oBrowser, oFso, sIEVersion, iIEVersion, iCnt

	Services.StartTransaction "OpenBrowserInstance" ' Timer Begin
	Reporter.ReportEvent micDone, "OpenBrowserInstance", "begin"
	
	' Check to verify passed parameters that they are not null or an empty string
	If IsNull(sBrwInstance) or sBrwInstance = "" Then
		Reporter.ReportEvent micFail, "Invalid Parameters", "Invalid parameter were passed to the OpenBrowserInstance function check passed parameters"
		OpenBrowserInstance = False ' Return Value
		Services.EndTransaction "OpenBrowserInstance" ' Timer End
		Exit Function
	End If
   
	' Launch Internet Explorer
	SystemUtil.Run("IEXPLORE.EXE")

	' Create Browser Object
	Set oAppBrowserInst = Description.Create()
	oAppBrowserInst("MicClass").Value = "Browser"
	oAppBrowserInst("openedbytestingtool").Value = True
	oAppBrowserInst("CreationTime").Value = sBrwInstance
	'oAppBrowserInst("title").Value = "about:blank" 'Environment.Value("BROWSER_TITLE")

	windowID = Browser(oAppBrowserInst).GetROProperty("hwnd") ' Get the window id for newly created browser window
	Environment("WINDOW_ID") = windowID ' Set Environment Variable to Browser Window ID use to determine correct window

	' Navigate to about:blank
	Browser(oAppBrowserInst).Navigate "about:blank"
	Browser(oAppBrowserInst).Highlight

	' Verify IE Window Exists, if it does navigate to application URL
	If Window("Hwnd:=" & windowID).Exist(3) Then
	   Reporter.ReportEvent micPass, "Launching Browser Window", "The Browser window was successfully opened."
	   sUrl = Environment("TEST_BROWSER_URL")
	   Browser(oAppBrowserInst).Navigate sUrl
	   Browser(oAppBrowserInst).Highlight
	   
	   ' Handle the Certificate Error Page in Stage Environment -- Added by Gov on 01/05/2012
		If Instr(1, Environment("TEST_BROWSER_URL"), "stage", 1) > 0 Then
			Set oFso = CreateObject("Scripting.FileSystemObject") 
			sIEVersion = oFso.GetFileVersion("C:\Program Files\Internet Explorer\iexplore.exe") 
			iIEVersion=CInt(Left(sIEVersion,1))
			iCnt = 0
			If iIEVersion >= 7 Then
				If Browser("name:=Certificate.*").Exist(2) Then
				' If the Page is taking more time to load the Certificate Error Code is written to handle the Max Time of 60 Seconds
					Do
						Browser("name:=Certificate.*").Page("title:=Certificate.*").Link("name:= Continue.*").Click
						iCnt = iCnt + 1
					Loop Until (Browser("name:=Certificate.*").Exist(1) = False OR iCnt = 60)
				End If
			Else
				If Browser("Hwnd:=" & Environment.Value("WINDOW_ID")).Dialog("regexpwndtitle:=Security Alert").Exist(2) Then
					Do
						Browser("Hwnd:=" & Environment.Value("WINDOW_ID")).Dialog("regexpwndtitle:=Security Alert").WinButton("text:=&Yes").Click
						iCnt = iCnt + 1
					Loop Until (Browser("Hwnd:=" & Environment.Value("WINDOW_ID")).Dialog("regexpwndtitle:=Security Alert").Exist(1) = False OR iCnt = 60)
				End If
			End If	
			Set oFso=Nothing
		End If

	   'Call init_securityAlert() ' NO NEED - Calls local function to check for Security Alert Dialog Box
	   Browser("Hwnd:=" & Environment.Value("WINDOW_ID")).Sync
	   Window("Hwnd:=" & Environment.Value("WINDOW_ID")).Maximize
	   oAppBrowserInst("hwnd").Value = Environment.Value("WINDOW_ID")
	   Environment("BROWSER_OBJ_INSTANCE") = oAppBrowserInst ' Set Environment Variable to Browser Object
	   OpenBrowserInstance = True ' Return Value
	   
	Else ' IE Window Does Not Exist
	   Reporter.ReportEvent micFail, "Launching Browser Window", "The Browser window does not exist after opening it."
	   OpenBrowserInstance = False ' Return Value
	End If
						
	' Clear Object Variables
	Set oAppBrowserInst = Nothing

	Reporter.ReportEvent micDone, "OpenBrowserInstance", "End"
	Services.EndTransaction "OpenBrowserInstance" ' Timer End
	End Function
End Class

'**********************************************************************************************
'*                            Class Instantiation                                         
'**********************************************************************************************
dim lin

set lin = new Login