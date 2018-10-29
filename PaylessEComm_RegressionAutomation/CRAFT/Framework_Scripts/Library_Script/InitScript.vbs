Class InitScript

	'=================================================================================
	' Purpose		: Function to close all open browsers before execution starts
	' Author		: Cognizant Tecnology Solutions
	' Reviewer		:
	'=================================================================================
	Function CloseAllBrowsers()

		Systemutil.CloseProcessByName"IEXPLORE.EXE"

	End Function

	'=================================================================================
	' Purpose		: Function for Test data are logically organized for end to end functionality 
	'			  testing for entire business / transaction cycle. In this concept read only 
	'			  external test data is imported in to run time data tables. Using scenario 
	'			  specific global sheets the run time data is updated dynamically where ever 
	'			  required. 
	' Author		: Cognizant Tecnology Solutions
	' Reviewer		:
	'=================================================================================

	Function ImportData(ByVal strPath, ByVal StrSheetName)
	
		DataTable.AddSheet StrSheetName
		DataTable.ImportSheet strPath,"Global", StrSheetName

	End Function

	'=================================================================================
	'Function Name	:Invoke Application Under Test
	'=================================================================================

	Function InvokeAUT(ByRef InitEnvironmentVariables)
	
		Dim strUrlPath 
			
		strUrlPath = InitEnvironmentVariables.URLPath
	   
		SystemUtil.Run "iexplore", strUrlPath
		
	End Function


End Class
