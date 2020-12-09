'------------------------------------------------------------------
'--Script name - Capture objects from script created from robot framework
'--Author  - Karthik Kumar
'--Date 	  - 31-Oct-18
'--Description - To capture the object names from the script generated from robot and save it to the excel sheet in UAIF format
'------------------------------------------------------------------
Func_ExcelProcessKill ("EXCEL.EXE")
Set objfile = CreateObject("Scripting.FileSystemObject")
'---Get the base folder for the script
basefolder=objfile.GetParentFolderName(WScript.ScriptFullName)
'basefolder="C:\backup\Work\GTA\Macro utilities\Robot integraiton"
curPath = basefolder & "\Scripts"
'msgbox objfile.GetParentFolderName(WScript.ScriptFullName)
Set base = objfile.GetFolder(curPath)
Set files = base.Files
Set ExlApp = CreateObject("Excel.Application")
ExlApp.Visible=True
'--Open the config sheet
configpath=	basefolder & "\Config.xlsx"
Set configworkbook=ExlApp.Workbooks.Open(configpath)
Set Configsheet=configworkbook.WorkSheets("Config")
Set Configmapsheet=configworkbook.WorkSheets("Script-ActionMapping")
noofrows=Configmapsheet.usedrange.rows.count
'--Declare an array for storing the mapping values
ReDim arrmapvalues(noofrows-1,6)
For i = 1 To noofrows
	irowarrindex=i-1
	For j = 1 To 7
		icolarrindex=j-1
		arrmapvalues(irowarrindex,icolarrindex)=Configmapsheet.rows(i).columns(j).value
	Next
	
Next
'--
iscriptrownum=2
itestcasenum=1
imastersheetnum=2
'--
Objectfilename=Configsheet.rows(2).columns(1).value
Objectsheetname=Configsheet.rows(2).columns(2).value
Testdatafilename=Configsheet.rows(2).columns(3).value
Testdatafilesheet=Configsheet.rows(2).columns(4).value
Scriptfilesheet=Configsheet.rows(2).columns(5).value
Testcasename=Configsheet.rows(2).columns(6).value
Mastersheetname=Configsheet.rows(2).columns(7).value
objectsheetpath = basefolder & "\" & Objectfilename
mastersheetpath=basefolder & "\" & Mastersheetname
testdatapath=basefolder & "\" & Testdatafilename
configworkbook.Close
'--check if object sheet exists
If objfile.FileExists(objectsheetpath) Then
Else
	'Call function to create the object sheet along with script sheet
	 Call CreateObjectandscriptsheet(objectsheetpath,Objectsheetname,Scriptfilesheet)	
End If
		
'--Open the object sheet to add object details
Set objectworkbook=ExlApp.Workbooks.Open(objectsheetpath)
Set objectsheet=objectworkbook.WorkSheets(Objectsheetname)
Set objscriptsheet=objectworkbook.WorkSheets(Scriptfilesheet)
'--check if master sheet exists
If objfile.FileExists(mastersheetpath) Then
Else
	'Call function to create the master sheet 
	 Call CreateMastersheet(mastersheetpath,"Master")
End If
'--Open the mastersheet to add test case names and test data file name
Set objmaster=ExlApp.Workbooks.Open(mastersheetpath)
Set objmastersheet=objmaster.WorkSheets("Master")
Call Retrieveblankrowtestdatasheet(objmastersheet,irowno)
imastersheetnum=irowno
'--Open the test datasheet
'--check if test data sheet exists
If objfile.FileExists(testdatapath) Then
Else
	'Call function to create the test data sheet
	 Call CreateTestdatasheet(testdatapath,Testdatafilesheet)	
End If
Set otestdataworkbook=ExlApp.Workbooks.Open(testdatapath)
Set otestdatasheet=otestdataworkbook.Worksheets(Testdatafilesheet)
Call Retrieveblankrowtestdatasheet(otestdatasheet,irowno)
itestdatarowval=irowno

'--Get the used row count in script sheet
iscriptrowInt=objscriptsheet.usedrange.rows.count
iscriptrowInt=iscriptrowInt+1
'--Get the existing test case number for an test case in script sheet
Call Retrievemaxtestcasenumber(objscriptsheet,Testcasename,1,iusedtestcaseno)
itestcasenum=iusedtestcaseno+1

For Each File in files
	If LCase(objfile.GetExtensionName(File.Path) = "robot") Then
		sacttestcasename=Testcasename&" - "&itestcasenum		
		
		'--Open the robot script in robot file and convert to excel
		robotscriptpath=File.Path
		arrScriptName=split(robotscriptpath,"\")
		robotScriptName=arrScriptName(ubound(arrScriptName))
		robotScriptNamewotxt=Replace(robotScriptName,".robot","")
		'ExlApp.Workbooks.Open robotscriptpath,,,6,,,,,","
		ExlApp.Workbooks.Open robotscriptpath,,,6,,,,,"	" 'open with tab Delimited
		robotscriptexcelpath=curPath & "\" & robotScriptNamewotxt & ".xls"
		If objfile.FileExists(robotscriptexcelpath) Then
			objfile.DeleteFile (robotscriptexcelpath)
        End If
		ExlApp.ActiveWorkbook.SaveAs (robotscriptexcelpath)
		ExlApp.ActiveWorkbook.Close
		Set robotworkbook=ExlApp.Workbooks.Open(robotscriptexcelpath)
		Set robotexcelworksheet=robotworkbook.WorkSheets(robotScriptNamewotxt)		
		'--Find the row number containing test cases
		For i = 1 To robotexcelworksheet.usedrange.rows.count
			srowvalue=robotexcelworksheet.rows(i).columns(1).value
			If Instr(1,srowvalue,"*** Test Cases ***")>0	 Then
				itestcasestart=i+2
				Exit For
			End If
		Next
		'----Check if any row contains multiple commands and insert it to an new line
		
		iobjusedrows=objectsheet.usedrange.rows.count
		For j = itestcasestart To robotexcelworksheet.usedrange.rows.count
			srowsecondcommand=""
			srowvalue=Trim(robotexcelworksheet.rows(j).columns(1).value)
			If srowvalue<>"" and Instr(1,srowvalue,"    ")>0 Then
				srowvaluesplit=Split(srowvalue,"    ")
				If Ubound(srowvaluesplit)=3  Then' If there are multiple commands in same line	
					If Instr(srowvaluesplit(3),"//")>0 Then
						srowfirstcommand=	srowvaluesplit(0) & srowvaluesplit(1)
						srowsecondcommand=	srowvaluesplit(2) & srowvaluesplit(3)
						scellref="A"&j+1&":"&"A"&j+1
						robotexcelworksheet.Range(scellref).EntireRow.Insert
						robotexcelworksheet.rows(j).columns(1).value=srowfirstcommand
						robotexcelworksheet.rows(j+1).columns(1).value=srowsecondcommand
						'j=j+2
					End If					
				End If
				
			End If
		Next
		robotworkbook.Save
		
		
		'---------
		iobjusedrows=objectsheet.usedrange.rows.count
		
		
		
		istepno=1
		'--For all the test cases step, capture the object name
		For i = itestcasestart To robotexcelworksheet.usedrange.rows.count
			testdataval=""
			iobjrowval=0
			xpathactval =""
			smanualstep=""
			sexpectedoutcome=""
			saction=""
			sobjectreq=""
			sobjecttype=""
			slocatortype=""
			surlval=""
			slocatorvalstart=""
			slocatorvalend=""
			slocatorlen=""
			sattribute=""
			slocatorval=""
			srowvalue=Trim(robotexcelworksheet.rows(i).columns(1).value)
			srowvalue=Replace(srowvalue,"	","")
			
			Call Retrievemappedsteps(arrmapvalues,srowvalue,smanualstep,sexpectedoutcome,saction,sobjectreq,sobjecttype)
			If sobjectreq = "Yes" and srowvalue <> "" Then
				
				'If the test case step has xpath along with index value
				If Instr(1,srowvalue,"//")>0 And Instr(1,srowvalue,"xpath=")>0 Then
					Call Retrievexpathparameterswithindex(srowvalue,xpathactval,testdataval,sattribute)
					
					
				'--If the test case step has xpath without index
				ElseIf Instr(1,srowvalue,"//")>0 And Instr(1,srowvalue,"xpath=")= 0 Then 
					Call Retrievexpathparameters(srowvalue,xpathactval,testdataval,sattribute,slocatorval)
				End If
				If xpathactval <> "" Then
					Call FillObjectsheet(objectsheet,iobjrowval,xpathactval,slocatorval,sattribute,slocatortype)
					
					'--If test data to be added
					If testdataval <> "" Then
						Call Verifyexistingcoltestdatasheet(otestdatasheet,sattribute,icolno,bresult)
						'If test data already exists
						If bresult=True Then
							
							otestdatasheet.rows(itestdatarowval).columns(icolno).value=testdataval
						Else
							Call Retrieveblanktestdatasheet(otestdatasheet,icolno)						
							otestdatasheet.rows(1).columns(icolno).value=sattribute
							otestdatasheet.rows(itestdatarowval).columns(icolno).value=testdataval
						End If
					End If
				End If
			End If
			'-Check if line is not blank
			If srowvalue <> "" Then
				'--If keyword is open browser, get the url value
				If Instr(1,srowvalue,"Open Browser")>0 Then
					iopenbstart=Instr(1,srowvalue,"Open Browser")
					openbstring=Mid(srowvalue,iopenbstart)
					sbrowsersplit=Split(openbstring,"    ")
					surlval=sbrowsersplit(1)
				End If
				'--Update the script sheet with all required values
				objscriptsheet.rows(iscriptrowInt).columns(1).value =sacttestcasename  'Test case name
				objscriptsheet.rows(iscriptrowInt).columns(3).value ="Step "&istepno  'Step name with number
				'-----Manual step
				If Instr(1,smanualstep,"[Variable Name]")>0 Then
					smanualstep=Replace(smanualstep,"[Variable Name]",sattribute)
					
				End If
				If Instr(1,smanualstep,"[Variable Value]")>0 Then
					smanualstep=Replace(smanualstep,"[Variable Value]",testdataval)
				End If
				If Instr(1,smanualstep,"[Test data]")>0 Then
					smanualstep=Replace(smanualstep,"[Test data]",testdataval)
				End If
				objscriptsheet.rows(iscriptrowInt).columns(4).value =  smanualstep'Test step name
				'--Expected result
				If Instr(1,sexpectedoutcome,"[Variable Name]")>0 Then
					sexpectedoutcome=Replace(sexpectedoutcome,"[Variable Name]",sattribute)
				End If
				If Instr(1,sexpectedoutcome,"[Variable Value]")>0 Then
					sexpectedoutcome=Replace(sexpectedoutcome,"[Variable Value]",testdataval)
				End If
				If Instr(1,sexpectedoutcome,"[Test data]")>0 Then
					sexpectedoutcome=Replace(sexpectedoutcome,"[Test data]",testdataval)
				End If
				objscriptsheet.rows(iscriptrowInt).columns(5).value =  sexpectedoutcome'Expected result
				'--
				objscriptsheet.rows(iscriptrowInt).columns(6).value =  sobjecttype'object type
				objscriptsheet.rows(iscriptrowInt).columns(7).value =  saction'UAIF keyword
				If sobjectreq = "Yes" Then
					objscriptsheet.rows(iscriptrowInt).columns(8).value =  sattribute'Object name	
					If testdataval <> "" Then
						objscriptsheet.rows(iscriptrowInt).columns(9).value =  Testdatafilesheet'Test data sheet name	
						objscriptsheet.rows(iscriptrowInt).columns(10).value =  sattribute'Test data column name	
					End If
				objscriptsheet.rows(iscriptrowInt).columns(13).value =  "Yes" 'Screenshot column value
				ElseIf Instr(1,srowvalue,"Open Browser")>0 Then
					objscriptsheet.rows(iscriptrowInt).columns(10).value =  surlval'url value
				End If
				
				istepno=istepno+1
				iscriptrowInt=iscriptrowInt+1
			End If ' If srowvalue <> ""
		Next
		'--Write values to the master sheet
		objmastersheet.rows(imastersheetnum).columns(1).value=imastersheetnum-1 ' SI no
		objmastersheet.rows(imastersheetnum).columns(4).value="NA" 'Module Key
		objmastersheet.rows(imastersheetnum).columns(5).value=sacttestcasename ' Testcase name
		objmastersheet.rows(imastersheetnum).columns(6).value="NA" 'Testcase Key
		objmastersheet.rows(imastersheetnum).columns(7).value="Yes" ' Executable
		objmastersheet.rows(imastersheetnum).columns(8).value=Objectfilename ' Script excel name
		objmastersheet.rows(imastersheetnum).columns(9).value=Scriptfilesheet ' Test case sheet name
		objmastersheet.rows(imastersheetnum).columns(10).value="Chrome" 'Browser
		objmastersheet.rows(imastersheetnum).columns(11).value= itestdatarowval-1 'From Inputrownum
		objmastersheet.rows(imastersheetnum).columns(12).value= 1 'Iteration
		objmastersheet.rows(imastersheetnum).columns(13).value= Testdatafilename 'Test data excel name
		objmastersheet.rows(imastersheetnum).columns(14).value= Testdatafilesheet 'Input sheet
		objmastersheet.rows(imastersheetnum).columns(15).value= "Output" 'Output sheet
		objmastersheet.rows(imastersheetnum).columns(16).value= Objectfilename 'Object repository excel name
		objmastersheet.rows(imastersheetnum).columns(17).value= Objectsheetname 'Object repository sheet name
		robotworkbook.Close
		itestcaseInt=itestcaseInt+1
		'iscriptrowInt=iscriptrowInt+1
		itestdatarowval=itestdatarowval+1
		itestcasenum=itestcasenum+1
		imastersheetnum=imastersheetnum+1
	End If
	
Next
objmaster.Save
objmaster.Close
otestdataworkbook.Save
otestdataworkbook.Close
objectworkbook.Save
objectworkbook.Close
Msgbox "Script steps and object details has been added to sheet successfully!!!"
ExlApp.Quit
Set ExlApp=Nothing
Public Function Func_ExcelProcessKill (strProcess)
	Dim FunctionName: FunctionName = UCase ("Func_ProcessKill")
	Dim strPreFix: strPreFix = "In " & pcThisLib & " - " & FunctionName & ","
	On Error Resume Next
	aProcess = Split (strProcess, "|", -1, 1)
	If isArray(aProcess) Then
		For Each ProcessItem In aProcess ' Now loop and find process to stop.
			If Func_ProcessExist (ProcessItem) = micPass Then ' If exist just stop it.
				Func_ProcessKill="Open"
				Reporter.ReportEvent micDone, "Func_ProcessKill", strPreFix & " stopped '" & ProcessItem & "' process."
				SystemUtil.CloseProcessByName(ProcessItem)'ProcessItem)
			End If
		Next
	End If
	' Restore default error handling.
	On Error Goto 0
End Function

Public Function Func_ProcessExist (ByVal strProcess)
	Dim FunctionName: FunctionName = UCase ("Func_ProcessExist")
	Dim strPreFix: strPreFix = "In " & pcThisLib & " - " & FunctionName & ","
	Set objSWbemServices = GetObject("winmgmts:")
	Set colSWbemObjectSet = objSWbemServices.ExecQuery("SELECT * FROM Win32_Process Where Name = '"& strProcess &"'")
	If colSWbemObjectSet.Count > 0 Then
		Func_ProcessExist = micPass ' Found the process.
	Else
		Func_ProcessExist = micFail ' Did not find the process.
	End If
	Set objSWbemServices = Nothing ' Clear object from memory.
	Set colSWbemObjectSet = Nothing ' Clear object from memory.
End Function

Public Function RetrieveblankrowObjsheet(objsheet,irowno)
For j = 2 To objsheet.usedrange.rows.count
	sobjpropval=objsheet.rows(j).columns(3).value
	If sobjpropval = "" Then
		Exit For
	End If
	
Next
	irowno = j
End Function
Public Function Retrieveblankrowtestdatasheet(objsheet,irowno)

For j = 2 To objsheet.usedrange.rows.count
	stemp=""
	For k = 1 To objsheet.usedrange.columns.count
		sobjpropval=objsheet.rows(j).columns(k).value
		If sobjpropval <> "" Then
			stemp=sobjpropval
		End If
	Next
	If stemp = "" Then
		Exit For
	End If
Next
	irowno = j
End Function
Public Function Verifyexistingcoltestdatasheet(objsheet,scolname,icolno,bresult)
	bresult = False
	icolno = 0
	For k = 1 To objsheet.usedrange.columns.count
		sdatacolval=Trim(objsheet.rows(1).columns(k).value)
		If sdatacolval = scolname Then
			bresult=True
			icolno=k
			Exit For
		ElseIf k = objsheet.usedrange.columns.count Then
			bresult=False
		End If
		
	Next

End Function
Public Function Retrieveblanktestdatasheet(objsheet,icolno)
	icolno = 0
	bstatus=False
	k= 1
	Do While bstatus=False
		sdatacolval=Trim(objsheet.rows(1).columns(k).value)
		If sdatacolval="" Then
			icolno=k
			bstatus=True
			Exit Do
		End If
		k = k + 1
	Loop
	

End Function

Public Function Verifyexistingobjdatasheet(objsheet,sobjname,irowno,bresult)
	bresult = False
	irowno = 0
	For k = 2 To objsheet.usedrange.rows.count
		sobjnamesheet=Trim(objsheet.rows(k).columns(1).value)
		If sobjnamesheet=sobjname Then
			irowno=k
			bresult=True
			Exit For
		End If
		
	Next

End Function

Public Function Retrievemappedsteps(arrconfigmap,srobotkeyword,smanualstep,sexpectedoutcome,saction,sobjectreq,sobjecttype)
	bresult = False
	irowno = 0

	For k = 1 To Ubound(arrconfigmap)
		temp=Trim(arrconfigmap(k,1))
		If temp="" Then
			Exit For
		End If
		'-If expected keyword is available in mapping sheet
		If Instr(1,srobotkeyword,temp)>0 Then
			smanualstep=arrconfigmap(k,2)
			sexpectedoutcome=arrconfigmap(k,3)
			saction=arrconfigmap(k,4)
			sobjectreq=arrconfigmap(k,5)
			sobjecttype=arrconfigmap(k,6)
			bresult=True
			Exit For
		End If
		
	Next
	'--If expected keyword is not available in mapping sheet
	If bresult=False Then
		smanualstep="Robot keyword is not available in mapping sheet - "&srobotkeyword
		sexpectedoutcome="Robot keyword is not available in mapping sheet - "&srobotkeyword
		saction = "Robot keyword is not available in mapping sheet - "&srobotkeyword
	End If
End Function


Function ConvertToLetter(iCol)
   Dim iAlpha
   Dim iRemainder
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Function Changeextensiontotxt(objfile,oldext,newext)
	strExtension = objFile.Extension

    strExtension = Replace(strExtension, oldext, newext)
	
    strNewName = objFile.Drive & objFile.Path & objFile.FileName & "." & strExtension
	If objfile.FileExists(strNewName) Then
		objfile.DeleteFile (strNewName)
    End If
    errResult = objFile.Rename(strNewName)
	
End Function

Public Function Retrievemaxtestcasenumber(objsheet,testcasename,itestcasecol,iusedtestcaseno)
	bresult = False
	iusedtestcaseno=0
	
	For k = 2 To objsheet.usedrange.rows.count
		sacttestcasename=Trim(objsheet.rows(k).columns(itestcasecol).value)
		If Instr(1,sacttestcasename,testcasename)>0 Then
			sacttestcasenamesplit=Split(sacttestcasename,"-")
			iusedtestcaseno=Trim(sacttestcasenamesplit(1))
			bresult=True
			
		End If
		
	Next
	
End Function
Function CreateObjectandscriptsheet(sheetfilepath,objsheetname,scriptsheetname)
	Set objExcel = CreateObject("Excel.Application")
	Set objworkbook=objExcel.Workbooks.Add
	objworkbook.SaveAs(sheetfilepath)
	Set objfnobj=objworkbook.Sheets.Add
	objfnobj.Name=objsheetname
	Set objfnscriptsheet=objworkbook.Sheets.Add
	objfnscriptsheet.Name=scriptsheetname
	'--Add headers for object sheet
	scellref="A1:E1"
	With objfnobj
		.Cells(1,1).Value="OBJECT VARIABLE NAME"
		.Cells(1,2).Value="OBJECT_PROPERTY"
		.Cells(1,3).Value="PROPERTY_VALUE"
		.Cells(1,4).Value="OBJECTTYPE"
		.Range("A1").ColumnWidth=25
		.Range("B1").ColumnWidth=20
		.Range("C1").ColumnWidth=35
		.Range("D1").ColumnWidth=12
				
		.Range(scellref).Font.Bold=True	
	End With
	'--Add headers for script sheet
	scellref="A1:M1"
	With objfnscriptsheet
		.Cells(1,1).Value="TESTCASE NAME"
		.Cells(1,2).Value="PAGE NAME"
		.Cells(1,3).Value="STEP NO"
		.Cells(1,4).Value="TESTSTEP NAME"
		.Cells(1,5).Value="EXPECTED"
		.Cells(1,6).Value="SELECTOBJECTTYPE"
		.Cells(1,7).Value="ACTION"
		.Cells(1,8).Value="OBJECTNAME"
		.Cells(1,9).Value="INPUT EXCEL"
		.Cells(1,10).Value="INPUT DATA"
		.Cells(1,11).Value="TESTDATA EXCEL NAME"
		.Cells(1,12).Value="SHEET NAME"
		.Cells(1,13).Value="SCREENSHOT"
		.Range("A1").ColumnWidth=30
		.Range("B1").ColumnWidth=11
		.Range("C1").ColumnWidth=9
		.Range("D1").ColumnWidth=40
		.Range("E1").ColumnWidth=40
		.Range("F1").ColumnWidth=20
		.Range("G1").ColumnWidth=20
		.Range("H1").ColumnWidth=15
		.Range("I1").ColumnWidth=12
		.Range("J1").ColumnWidth=30
		.Range("K1").ColumnWidth=22
		.Range("L1").ColumnWidth=12
		.Range("M1").ColumnWidth=12
		.Range(scellref).Font.Bold=True	
	End With
	objworkbook.Save
	objworkbook.close
	objExcel.Quit

End Function

Function CreateTestdatasheet(sheetfilepath,objsheetname)
	Set objExcel = CreateObject("Excel.Application")
	Set objworkbook=objExcel.Workbooks.Add
	objworkbook.SaveAs(sheetfilepath)
	Set objfnobj=objworkbook.Sheets.Add
	objfnobj.Name=objsheetname
	
	objworkbook.Save
	objworkbook.close
	objExcel.Quit

End Function
Function CreateMastersheet(sheetfilepath,objsheetname)
	Set objExcel = CreateObject("Excel.Application")
	Set objworkbook=objExcel.Workbooks.Add
	objworkbook.SaveAs(sheetfilepath)
	Set objmaster=objworkbook.Sheets.Add
	objmaster.Name=objsheetname
	'--Add headers for master sheet
	scellref="A1:O1"
	With objmaster
		.Cells(1,1).Value="SL.NO"
		.Cells(1,2).Value="REQUIREMENT NAME"
		.Cells(1,3).Value="MODULE NAME"
		.Cells(1,4).Value="MODULE KEY"
		.Cells(1,5).Value="TESTCASE NAME"
		.Cells(1,6).Value="TESTCASE KEY"
		.Cells(1,7).Value= "EXECUTABLE"
		.Cells(1,8).Value="SCRIPT EXCEL NAME"
		.Cells(1,9).Value="TEST CASE SHEET NAME"
		.Cells(1,10).Value="BROWSER"
		.Cells(1,11).Value="FromInputRowNum"
		.Cells(1,12).Value="ITERATION"
		.Cells(1,13).Value="TESTDATA EXCEL NAME"
		.Cells(1,14).Value="INPUT SHEET"
		.Cells(1,15).Value="OUTPUT SHEET"
		.Cells(1,16).Value="OBJECT REPOSITORY EXCEL NAME"
		.Cells(1,17).Value="OBJECT REPOSITORY SHEET NAME"
		.Range("A1").ColumnWidth=7
		.Range("B1").ColumnWidth=20
		.Range("C1").ColumnWidth=20
		.Range("D1").ColumnWidth=30
		.Range("E1").ColumnWidth=13
		.Range("F1").ColumnWidth=21
		.Range("G1").ColumnWidth=25
		.Range("H1").ColumnWidth=10
		.Range("I1").ColumnWidth=17
		.Range("J1").ColumnWidth=11
		.Range("K1").ColumnWidth=25
		.Range("L1").ColumnWidth=13
		.Range("M1").ColumnWidth=15
		.Range("N1").ColumnWidth=20
		.Range("O1").ColumnWidth=15
		.Range("P1").ColumnWidth=35
		.Range("Q1").ColumnWidth=35
		.Range("R1").ColumnWidth=35
		.Range(scellref).Font.Bold=True	
	End With
	objworkbook.Save
	objworkbook.close
	objExcel.Quit

End Function

Function Retrievexpathparameterswithindex(objpropvalue,xpathactval,testdataval,sattribute)
	testdataval=""
	slocatortype="XPath"
	ixpathstart=Instr(1,objpropvalue,"(")
	ixpathend=Instr(1,objpropvalue,")")
	ixpathlen=ixpathend-ixpathstart
	xpathwithinbracket=Mid(srowvalue,ixpathstart,ixpathlen)
	xpathstr=Mid(srowvalue,ixpathstart)
	indexstr=Mid(srowvalue,ixpathend+1)					
	iindexvaluestart=Instr(1,indexstr,"[")+1
	iindexvalueend=Instr(1,indexstr,"]")
	iindexlen=iindexvalueend-iindexvaluestart
	indexval=Mid(indexstr,iindexvaluestart,iindexlen)			
	'xpathactval=xpathstr & "|" & "index:=" & indexval
	If Instr(1,xpathstr,"    ")>0 Then
		xpathsplit=Split(xpathstr,"    ")
		xpathactval=xpathsplit(0)
		testdataval=xpathsplit(1)
		
	Else
		xpathactval=xpathstr
	End If
	'--Get the attribute name
	If Instr(1,xpathactval,"=")>0 Then
		iattributestart=Instr(1,xpathactval,"=")+2
		iattributeend=Instr(1,xpathactval,"]")-1
		iattributelen=iattributeend-iattributestart
		attribute=Mid(xpathstr,iattributestart,iattributelen)
		sattribute=attribute&indexval
	Else
		iattributestart=4
		attribute=Mid(xpathwithinbracket,iattributestart)						
		sattribute=attribute&indexval
	End If

End Function

Function Retrievexpathparameters(objpropvalue,xpathactval,testdataval,sattribute,slocatorval)
	slocatortype="XPath"
	ixpathstart=Instr(1,objpropvalue,"//")
	xpathval=Mid(objpropvalue,ixpathstart)
	If Instr(1,xpathval,"    ")>0 Then
		xpathsplit=Split(xpathval,"    ")
		xpathactval=xpathsplit(0)
		testdataval=xpathsplit(1)
	Else
		xpathactval=xpathval
	End If
	'--Get the attribute name
	If Instr(1,xpathactval,"=")>0 Then
		'--Check if the keyword contains name locator type
		If Instr(1,xpathactval,"name=")>0 Then
			slocatortype="Name"	
			slocatorvalstart=Instr(1,xpathactval,"name=")+6
			slocatorvalend=Instr(1,xpathactval,"]")-1
			slocatorlen=slocatorvalend-slocatorvalstart
			slocatorval=Mid(xpathactval,slocatorvalstart,slocatorlen)
		'--Check if the keyword contains ID locator type	
		ElseIf Instr(1,xpathactval,"id=")>0 Then
			slocatortype="ID"	
			slocatorvalstart=Instr(1,xpathactval,"id=")+4
			slocatorvalend=Instr(1,xpathactval,"]")-1
			slocatorlen=slocatorvalend-slocatorvalstart
			slocatorval=Mid(xpathactval,slocatorvalstart,slocatorlen)
'		'--Check if the keyword contains className locator type
'		ElseIf Instr(1,xpathactval,"class=")>0 Then
'			slocatortype="Class Name"	
'			slocatorvalstart=Instr(1,xpathactval,"class=")+7
'			slocatorvalend=Instr(1,xpathactval,"]")-1
'			slocatorlen=slocatorvalend-slocatorvalstart
'			slocatorval=Mid(xpathactval,slocatorvalstart,slocatorlen)
		Else
			slocatortype="XPath"	
		End If
		If slocatortype="XPath" Then
			iattributestart=Instr(1,xpathactval,"=")+2
			iattributeend=Instr(1,xpathactval,"]")-1
			iattributelen=iattributeend-iattributestart
			sattribute=Mid(xpathactval,iattributestart,iattributelen)
		Else
			sattribute=slocatorval
		End If
	ElseIf Instr(1,xpathactval,"contains")>0 Then
		xpathsplit=Split(xpathactval,",")
		sattributeraw=Trim(xpathsplit(1))
		sattributeraw=Replace(sattributeraw,"'","")
		sattributeraw=Replace(sattributeraw,"""","")
		icloseparen=Instr(sattributeraw,")")
		sattribute=Mid(sattributeraw,1,icloseparen-1)	
	Else
		iattributestart=3
		sattribute=Mid(xpathactval,iattributestart)						
		
	End If

End Function

Function FillObjectsheet(objectsheet,iobjrowval,xpathactval,slocatorval,sattribute,slocatortype)
	Call Verifyexistingobjdatasheet(objectsheet,sattribute,irowno,bresult)
	If bresult=True Then
		iobjrowval=irowno
	Else
		Call RetrieveblankrowObjsheet(objectsheet,irowno)
		iobjrowval=irowno
	End If
	'--Set the object property value
	If slocatortype="XPath" Then 
		objectsheet.rows(iobjrowval).columns(3).value=xpathactval
	Else
		objectsheet.rows(iobjrowval).columns(3).value=slocatorval
	End If
	
	objectsheet.rows(iobjrowval).columns(2).value=slocatortype
	objectsheet.rows(iobjrowval).columns(4).value="Element"	
	objectsheet.rows(iobjrowval).columns(1).value=sattribute
	
End Function

