'**********************************************************
'ajaxSync.vbs 
'
'   ajaxSyncRequest  	'ajax-based Processing Request/Please Wait wheel
'   pageObjsCount		'return a count of all child objects currently in an ajax-based Browser page
'   ajaxBrowserSync     'ajax-base Browser page sync. w/timeout
'
'Listing of available functions: See function header for usage. 
'Please list in ALPHABETICAL for the added function name with a short description.
'**********************************************************

'**********************************************************
Option Explicit 
Class ajaxSyncFunctions
	Function ajaxSyncRequest(sElementName,iTimeout)
		'**************************************************************************************************************************************************************
		'Purpose: Wait for the ajax-based (WebElement processing request wheel)  object specified to finish processing a request.
		'Parameters: sElementName = string - the WebElement object innertext value
		'                         iTimeout = numeric - seconds 
		'Requires: Prior to calling this function:
		'                     a. A process must be requested
		'                     b. The AJAX 'Please Wait' or 'Processing Request' wheel object is expected to appear.
		'Calls: None
		'Returns: True/False
		'Usage: Call ajaxSync.ajaxSyncRequest("Processing Request",60)	'waiting for the 'Processing Request wheel object to finish in 60 secs time out
		'       Call ajaxSync.ajaxSyncRequest("Please Wait",60)	        'waiting for the 'Please Wait' wheel object to finish in 60 secs time out
		'Created by: Hung Nguyen 6/24/11
		'Modified: Hung Nguyen 9/6/12 Modified function used from InSite proj. for ChartSync use.
		'                   Hung Nguyen 10/15/12 - Added to handle error if exists
		'**************************************************************************************************************************************************************
		services.StartTransaction "ajaxSyncRequest"
		Dim oAppBrowser,cnt,oElement,oChildren,iChildrenCount,i,iFound,iX
		ajaxSyncRequest=False	'init return value

		'verify parameters
		if not isnumeric(iTimeout) then
			reporter.ReportEvent micFail,"ajaxSyncRequest","Parameter 'iTimeout' must contains a numeric value."
			Exit Function
		End If
		if sElementName="" or isempty(sElementName) then
			reporter.ReportEvent micFail,"ajaxSyncRequest","Parameter 'sElementName' must contains a value."
			Exit Function
		End If

		'any Browser obj which contains the ajax-based processing request obj - expect one only!
		Set oAppBrowser=Browser(Environment("BROWSER_OBJ")).Page(Environment("BROWSER_OBJ"))
		oAppBrowser.Sync

		On error resume next
		cnt=0
		Do while cnt <=iTimeout
			Set oElement=description.Create    
			oElement("micclass").value="WebElement"
			oElement("class").value="rich-mpnl-body"
			oElement("html tag").value="TD"
			oElement("innertext").value=sElementName

			Set oChildren=oAppBrowser.ChildObjects(oElement)
			iChildrenCount=oChildren.Count
			reporter.ReportEvent micInfo,"object found=" &iChildrenCount &". Checking if visible..."  ,""

			If iChildrenCount > 0 Then			
				For i=0 to iChildrenCount - 1
					iFound=0
					If oChildren(i).GetROProperty("x") > 0 Then
						iX=oChildren(i).GetROProperty("x")
						reporter.ReportEvent micInfo,"Object index '" &i &"' is visible at position=" &iX,"Expect to disappear at next Do loop..."
						iFound=1
						Exit For
					Else
						reporter.ReportEvent micInfo,"Object index '" &i &"' is not visible",""
					End If
				Next	'child

				If iFound=0 Then 
					ajaxSyncRequest=True	'return value
					reporter.ReportEvent micInfo,"Object '" &sElementName &"' does not exist or no longer visible.",""
					Exit Do
				End If 
			Else
				ajaxSyncRequest=True	'return value
				reporter.ReportEvent micInfo,"Object '" &sElementName &"' not found",""
				Exit Do
			End If

			Set oElement=Nothing
			Set oChildren=Nothing
			reporter.ReportEvent micInfo,"next do loop... til object disappears.",""
			Wait(1)		'sec loop interval
			cnt=cnt+1
		Loop

		On error goto 0	'reset
		Set oAppBrowser=Nothing
		services.EndTransaction "ajaxSyncRequest"
	End Function
		
	Function pageObjsCount(sBrowserTitle)
		'****************************************************************************************************
		'Purpose: Returns a collection/count of elements contained by an ajax-based web app.
		'         The object count will be used for Browser synchronization.
		'Parameter: None
		'Calls: None
		'Returns: Numeric - count of child element objects OR 0 if fails. 
		'Usage: iObjectCount=ajaxSync.pageObjsCount("Optum ChartSync")
		'Created by: Hung Nguyen 9/10/12
		'Modified:
		'****************************************************************************************************
		services.StartTransaction "pageObjsCount"
		pageObjsCount = 0  'init return value	
		Dim oElements
		
		'If Browser("title:=" &sBrowserTitle).Page("title:=" &sBrowserTitle).Exist(2) Then 
		If Browser(Environment("BROWSER_OBJ")).Page("title:=.*").Exist(2) Then 
			'Set oElements=Browser("title:=" &sBrowserTitle).Page("title:=" &sBrowserTitle).Object.all
			Set oElements=Browser(Environment("BROWSER_OBJ")).Page("title:=.*").Object.all
			pageObjsCount=oElements.length	'return value
			
			reporter.ReportEvent micInfo,"Total element objs contained by the Browser page obj = " &pageObjsCount,""
			Set oElements=Nothing 
		Else
			reporter.ReportEvent micFail,"Browser Title '" &sBrowserTitle &"' does not exist.",""
		End If 
		services.EndTransaction "pageObjsCount"
	End Function
	
	Sub ajaxBrowserSync(sBrowserTitle,iObjCountBefore,iTimeout)
		'****************************************************************************************************
		'Purpose: Ajax-based browser page synchronization.
		'         Use together w/pageObjCount()
		'Test Steps: 1. Call pageObjsCount() to obtain a before count of all objs contained by the Browser page obj.
		'            2. Click/select/navigate... an object on the Browser page
		'            3. Call this subroutine obtain an after count then compare the counts (before and after) 
		'               within a time out specified before move on to the next test step.
		'
		'NOTE: Use this in place of the QTP Wait() statement and or the Browser.sync method for an ajax-based web app.
		'      This function counts objects on page then compare count w/before to determine the Browser page synchronization.
		'      This function is different than the ajax-based Please Wait or Processing Request wheel synchronization function.
		'
		'Parameter: sBrowserTitle = the Browser title (property value)
		'           iObjCountBefore = numeric - total objects on the Browser page
		'           iTimeout = sec. waiting for the Browser page synchronization
		'           NOTE: time out must need to adjust accordingly to avoid waisting
		'Calls: ajaxSync.pageObjsCount
		'
		'Usage: iObjCntBefore=ajaxSync.pageObjsCount("Optum ChartSync")
		'       Click a tab submenu (or something...)
		'	    Call ajaxSync.BrowserSync("Optum ChartSync",iObjCntBefore,15)	'wait for page to sync w/15 secs timeout max.
		'       do something next...
		'
		'Returns: None
		'Created by: Hung Nguyen 9/10/12
		'Modified:
		'****************************************************************************************************
		services.StartTransaction "ajaxBrowserSync"
		Dim oElements,cnt,iSync,iCountAfter
		
		'verify parameters
		If Not IsNumeric(iObjCountBefore) Or Not IsNumeric(iTimeout) Then
			reporter.reportevent micFail,"Invalid parameter","obj count and time out values must be numeric."
			Exit Sub
		End If 
		If sBrowserTitle="" Or IsEmpty(sBrowserTitle) Then
			reporter.reportevent micFail,"Invalid parameter","Browser title can't be an empty string."
			Exit Sub
		End If
			
		'If Browser("title:=" &sBrowserTitle).Page("title:=" &sBrowserTitle).Exist(2) Then 		
		If Browser(Environment("BROWSER_OBJ")).Page("title:=.*").Exist(2) Then 		
			cnt=0
			iSync=0
			Do While cnt <= iTimeout
				Wait(1)
				cnt=cnt+1
				
				iCountAfter=pageObjsCount(sBrowserTitle)
				If iObjCountBefore <> iCountAfter Then
					iSync=1
					reporter.ReportEvent micInfo,"Count before=" &iObjCountBefore &vbnewline &"Count after=" &iCountAfter,"Page synchronization complete."
					Exit Do
				End If 			
			Loop 
			If iSync=0 Then
				reporter.ReportEvent micInfo,"Timed out - Count before=" &iObjCountBefore &vbnewline &"Count after=" &iCountAfter,"Page synchronization incomplete. NO change in counts."
			End If 
		Else
			reporter.ReportEvent micFail,"Browser Title '" &sBrowserTitle &"' does not exist.",""
		End If 
		
		Set oElements=Nothing 
		services.EndTransaction "ajaxBrowserSync"
	End Sub 

End Class
	
''***********Class instantiation************
Dim ajaxSync
Set ajaxSync = New ajaxSyncFunctions
''******************************************	