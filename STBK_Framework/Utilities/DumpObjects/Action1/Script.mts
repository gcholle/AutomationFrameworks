''----------------------------------------------------------------------' 
'Script Name: DumpObjects
'Creation Date: Dec 13, 2007
'Author: CASS Team
'Parameter Input: None
'Environment variable Dependencies: None
'Description: This script will read all the fields from the mainframe screen and dump that in to the specified MS Access table of the specified DataBase
' 
'Create the description object
Option Explicit
Dim DB,i
Dim oFieldDesc,oCollection
Dim DBCON
Dim oRst
Dim overwrite,table,Stmt,dbFailOnError, flag
Dim sure
Const adSchemaTables = 20
'Dim oFieldDesc,oCollection

Set oFieldDesc = Description.Create
oFieldDesc("micClass").Value = "TeField"
Set oCollection = TeWindow("TeWindow").TEScreen("TEScreen").ChildObjects(oFieldDesc)
	

'msgbox oCollection(i).row

Set DBCON = CreateObject("Adodb.connection")
	DBCON.Provider = "microsoft.jet.oledb.4.0"
	DB=InputBox("Please Enter Database path","Enter DB")
	i=1
	While DB="" AND  (i<>10)
		DB=InputBox( "Please Enter Database path","Enter DB")
		i=i+1
	Wend

	If i=10 Then
		ExitAction
		reporter.ReportEvent micFail,"DB Concection","Script execution Stopped: No Path Entered; Please enter path and try again"
	End If

	If DB<>"" and i<>10 Then
		DBCON.Properties("Data Source").Value = DB
    End If
	On error resume next
	DBCON.open
	
	If DBCON.state<>1 then ' Check if the connection is opened
			reporter.ReportEvent micFail,"DB Concection","Script execution Stopped: Failed to connect to database '"&DB&"'; Please check the path and try again"
			ExitAction
    end if

    table =InputBox( "Enter the Table Name")

	If table="" Then
		reporter.ReportEvent micFail,"Table initialization","No table entered; Script execution stopped"
		ExitAction
	end if 

	If IfTableExists(DBCON,table)=micPass Then
		overwrite=msgbox ("Table '"&table&"' already exists; Do you want to overwrite it?",vbYesNo)
	else
		
		Stmt="CREATE TABLE " &table &"(FieldName Text(50), FieldXY Text(50),StartRow Integer, StartColumn Integer, FieldLength Integer, Protected Text(50),DollarValue Text(1) )"
		
		DBCON.BeginTrans
		DBCON.execute Stmt,dbFailOnError
		DBCON.CommitTrans  
		reporter.ReportEvent micDone,"Table Created","New table created with table name '"&table&"'"
		
		StoreObjects table,DBCON,oCollection
    End If

	If overwrite=vbYes Then
					sure=msgbox ("Are you sure you want to overwrite the table"& table&"?",vbYesNo)
					If sure=vbyes Then
						Stmt="Delete * from "&table
						DBCON.BeginTrans
						DBCON.execute Stmt,dbFailOnError
						DBCON.CommitTrans  
						reporter.ReportEvent micDone,"Delete Records",dbFailOnError&" Records deleted from table "&"'table'"
						StoreObjects table,DBCON,oCollection
						
						DBCON.CommitTrans
					else
						ExitAction				
					End If
	End If
    
					
		

Public Function StoreObjects(Byval table, Byval DBCON, ByVal oCollection)
'	 	msgbox "in"
        Dim i,FieldName,ObjectEasyName,StartRow,StartColumn,Stmt
		Dim FieldXY
		Dim Protected
		Dim FieldLength
'		msgbox oCollection.count
			For i=0 to oCollection.count-1
					FieldName=oCollection(i).GetROProperty("attached text")
					FieldXY=table&"_"&oCollection(i).GetROProperty("start row")&"_"&oCollection(i).GetROProperty("start column")
					StartRow=oCollection(i).GetROProperty("start row")
					StartColumn=oCollection(i).GetROProperty("start column")
					Protected=oCollection(i).GetROProperty("protected")
					FieldLength=oCollection(i).GetROProperty("length")
					'Stmt = "insert into "table"(FieldName,FieldXY, StartRow, StartColumn,FieldLength,Protected)  VALUES ('" & FieldName & "', '" & FieldXY & "'," & StartRow & "," & StartColumn & "," & FieldLength &",'" & Protected & "')"	
					Stmt = "insert into "&table&" (FieldName,FieldXY, StartRow, StartColumn,FieldLength,Protected)  VALUES ('" & FieldName & "', '" & FieldXY & "'," & StartRow & "," & StartColumn & "," & FieldLength &",'" & Protected & "')"	
					'msgbox Stmt
						DBCON.BeginTrans
						DBCON.execute Stmt,dbFailOnError
						DBCON.CommitTrans
						
			Next				
		Reporter.ReportEvent micDone, "Table updated",i&" records added in table "&table
End Function


Function IfTableExists(Byval DBCON,Byval table)
			Dim RS,flag
			Set RS = DBCON.OpenSchema(adSchemaTables)
			Do Until RS.EOF
'--- Skip system tables
				If StrComp(RS("TABLE_TYPE").Value, "SYSTEM TABLE") <> 0 Then
'						msgbox RS("TABLE_NAME")
						If RS("TABLE_NAME").Value=table then
							IfTableExists=micPass
							msgbox RS("TABLE_NAME").Value
							flag=true
							Exit Do
						end If
                End If
				RS.movenext
            Loop
				If flag<>true Then
					 IfTableExists=micFail
				End If
						
End Function