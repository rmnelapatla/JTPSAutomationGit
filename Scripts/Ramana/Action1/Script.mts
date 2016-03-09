varTestpath =environment("TestDir")


environment.value("errorexist")="False"
environment.value("varpath") =Mid(varTestpath,1,Instrrev(Mid(varTestpath,1,instrrev(varTestpath,"\")-1),"\"))
environment.value("applicationtype")="Web"


'Add the Function libaries
LoadFunctionLibrary environment.value("varpath")&"FunctionLibrary\CommonFunctions_JTPS.txt"
LoadFunctionLibrary environment.value("varpath")&"FunctionLibrary\CommonFunctions.txt"



Call fn_ShowForm()

LoadFunctionLibrary environment.value("varpath")&"FunctionLibrary\Config - Sys.txt"
LoadFunctionLibrary environment.value("varpath")&"FunctionLibrary\Reportlib.txt"
LoadFunctionLibrary environment.value("varpath")&"FunctionLibrary\solReportlib.txt"

' Add the object repositry
'RepositoriesCollection.RemoveAll() 
'RepositoriesCollection.Add environment.value("ORregion")&"Portal.tsr"
'RepositoriesCollection.Add environment.value("ORregion")&"Trade.tsr"
RepositoriesCollection.Add environment.value("ORregion")&"JTPS.tsr"


If environment.value("strTestType") = "Regression_Testing" Then
	varDataTablePath=environment.value("datatableregion")&"Master_Ramana.xlsx"
	environment.value("varDataTabPath")=varDataTablePath
	invokeapplication environment.value("varpath")&"\Batfiles\KillExcel.bat"
	datatable.AddSheet "Master"
	datatable.ImportSheet vardatatablepath,"MasterSheet","Master"
	environment.value("varsheetname")= "Master"
ElseIf environment.value("strTestType") = "System_Testing" Then
	varDataTablePath=environment.value("datatableregion")&"Master_Ramana.xlsx"
	environment.value("varDataTabPath")=varDataTablePath
	invokeapplication environment.value("varpath")&"\Batfiles\KillExcel.bat"
	
	if environment.value("strsystemdata")<>"" then
		Call fn_processdata()
	End if
	datatable.AddSheet environment.value("strSytemDay")
	datatable.ImportSheet vardatatablepath,environment.value("strSytemDay"),environment.value("strSytemDay")
	environment.value("varsheetname")= environment.value("strSytemDay")
	
End If

Environment.value("Resultsheet")="Result_"&replace(replace(Now,"/","-"),":","-")

vTestcaserows=datatable.GetSheet(environment.value("varsheetname")).GetRowCount
Dim arrtestcaselen()
dim arrtestcaserep()
Dim arrtestcaseparameter()
j=0

For i = 1 To vTestcaserows
	datatable.Getsheet(environment.value("varsheetname")).SetCurrentRow i
	ReDim  preserve arrtestcaselen(j)
	Redim preserve arrtestcaserep(j)
	reDim Preserve arrtestcaseparameter(j)
	If datatable("Execution",environment.value("varsheetname"))="EOF" or datatable("TestScenarioName",environment.value("varsheetname"))="EOF" Then
		Exit for
	End If
	
	If datatable("Execution",environment.value("varsheetname"))="Y" Then
		arrtestcaselen(j)= datatable.Value("TestScenarioName",environment.value("varsheetname"))
		'arrtestcaseparameter(j)=datatable.Value("DataParameter",environment.value("varsheetname"))
		If datatable.Value("No_Of_Iterations",environment.value("varsheetname")) ="" Then
				 arrtestcaserep(j) = 0
			else
				arrtestcaserep(j) =cint( datatable.Value("No_Of_Iterations",environment.value("varsheetname")))
		End If
		arrtestcaseparameter(j)=datatable.Value("Parameter_TestData_SheetName",environment.value("varsheetname"))
		j=j+1
	End If
Next

For tc = 0 To Ubound(arrtestcaselen)-1
	
	varTestScenarioPath= environment.value("datatableregion")&arrtestcaselen(tc)&".xlsx"
	datatable.AddSheet  arrtestcaselen(tc)&"_Output"
	datatable.ImportSheet varTestScenarioPath,arrtestcaselen(tc),arrtestcaselen(tc)&"_Output"
	
	If arrtestcaserep(tc) >0 and arrtestcaseparameter(tc) <> "" Then
		Set vartestdataparxl = createobject("excel.application")
		vartestdataparxl.Workbooks.Open environment.value("datatableregion")&arrtestcaseparameter(tc)&".xlsx"
		vartestdataparxl.Application.Visible = False
		set vartestdataparsheet = vartestdataparxl.ActiveWorkbook.Worksheets("Sheet1")
		intstatusrowcount=vartestdataparsheet.UsedRange.Rows.Count
		
		For y=1 To intstatusrowcount
				If ucase(vartestdataparsheet.cells(y,1).value)=ucase(arrtestcaselen(tc)) Then
						vartestscrow= y
						Exit for
				End If
				Next
		vartestdataparxl.Workbooks(arrtestcaseparameter(tc)).Close
		Set vartestdataparxl=nothing
		Set vartestdataparsheet = nothing
	else
		arrtestcaserep(tc) = 0
		
	End If
' Getting rows count from datatable
	variteration =0
	do
		
		If arrtestcaserep(tc) >0 and arrtestcaseparameter(tc) <> "" Then 'to include the datadriven capability when more than one iteration happens				
			Set vartestdataparxl = createobject("excel.application")
			vartestdataparxl.Workbooks.Open environment.value("datatableregion")&arrtestcaseparameter(tc)&".xlsx"
			vartestdataparxl.Application.Visible = False
			set vartestdataparsheet = vartestdataparxl.ActiveWorkbook.Worksheets("Sheet1")
			intstatuscolcount=vartestdataparsheet.UsedRange.Columns.Count					
			for x=2 to intstatuscolcount 
				if vartestdataparsheet.cells(vartestscrow,x).value<>"" or lcase(trim(vartestdataparsheet.cells(vartestscrow,x).value))<>"empty" then
					vname=trim(vartestdataparsheet.cells(vartestscrow,x).value)
					vvalue=trim(vartestdataparsheet.cells(vartestscrow+1+variteration,x).value)
					If vname = "" or vvalue="" Then
						Exit for
					End If
					dict.add vname,vvalue
				End if
			Next
			vartestscenarioname=arrtestcaselen(tc)&"_"&(variteration+1)
			vartestdataparxl.Workbooks(arrtestcaseparameter(tc)).Close
			Set vartestdataparxl=nothing
			Set vartestdataparsheet = nothing
		else
			vartestscenarioname=arrtestcaselen(tc)
		End if
		
		
		intTestcaserows=datatable.GetSheet(arrtestcaselen(tc)&"_Output").GetRowCount
		
		If arrtestcaselen(tc)="EOF" Then
			Exit For
		End If
		
		intScriptstartval=1
		Dim slno
		For introw = 1 To intTestcaserows
			slno=slno+1
			datatable.GetSheet(arrtestcaselen(tc)&"_Output").SetCurrentRow(introw)
			boolRuntestcase=datatable("Run_TestCase",arrtestcaselen(tc)&"_Output")
			
			If ucase(boolRuntestcase)="EOF" then'  or environment.value("errorexist")="True"Then	
					'Datatable.ExportSheet environment.value("Results")&Environment.value("Resultsheet")&".xlsx",arrtestcaselen(tc)&"_Output"
					'Updatesheet arrtestcaselen(tc),Scenarioname,strFaillst,strInstIDreportExcel	
					'Updatesheet vartestscenarioname,Scenarioname,strFaillst,strInstIDreportExcel	
				Exit For
			End If
			
			If ucase(boolRuntestcase)="Y" Then
				boolen="Y"
			ElseIf ucase(boolRuntestcase)="N" Then
				boolen="N"
				For i= introw to intTestcaserows
					datatable.GetSheet(arrtestcaselen(tc)&"_Output").SetCurrentRow(i)
					boolRuntestcase=datatable("Run_TestCase",arrtestcaselen(tc)&"_Output")
					if boolRuntestcase="Y" then
							boolen="Y"
						introw=i
						Exit for
					ElseIf Ucase(boolRuntestcase)="EOF" Then
								Exit for ' to exit the current loop
								
					End if
					
					 Next	
				If Ucase(boolRuntestcase)="EOF" Then
							'Updatesheet vartestscenarioname,Scenarioname,strFaillst,strInstIDreportExcel	'update the status of the previous test case and close it
							Exit for 'to exit the current scenario
				End if	
			else
				boolen=boolen
			End If
				'Exit the test steps with N indicator
			boolRun=datatable("Run",arrtestcaselen(tc)&"_Output")
			
			If Ucase(boolRun)="N" Then
				For incrow=introw  To intTestcaserows
					datatable.GetSheet(arrtestcaselen(tc)&"_Output").SetCurrentRow(incrow)	
					If datatable("CaseID",arrtestcaselen(tc)&"_Output")<>"" Then
						strCaseDesc=datatable("CaseID",arrtestcaselen(tc)&"_Output")
						If Scenarioname="" Then
							Scenarioname=strCaseDesc
							intScriptstartval=intScriptstartval+1
						End If
						If strCaseDesc<>Scenarioname Then
							
							'Updatesheet vartestscenarioname,Scenarioname,strFaillst,strInstIDreportExcel	
							Scenarioname = strCaseDesc
							strFaillst=""
							intScriptstartval=intScriptstartval+1
						End If
					End If
					boolRun=datatable("Run",arrtestcaselen(tc)&"_Output")
					If boolRun = "Y" Then
						introw=incrow
						Exit For
					else
						intcrow=incrow+1
						
					End If	
					
				Next
				
				
			End If
			
			
			boolRun=datatable("Run",arrtestcaselen(tc)&"_Output")
			strAppType=datatable("Application_Type",arrtestcaselen(tc)&"_Output")
			strAction=datatable("Action",arrtestcaselen(tc)&"_Output")
			strBrwsID=datatable("BrwsrId",arrtestcaselen(tc)&"_Output")	
			strPageId=datatable("PageId",arrtestcaselen(tc)&"_Output")
			strFieldName=datatable("Field_Name",arrtestcaselen(tc)&"_Output")
			strInput=datatable("Input",arrtestcaselen(tc)&"_Output")
			intstepNum=datatable("Testcase_Step_Num",arrtestcaselen(tc)&"_Output")
			strcasedesdummy=datatable("CaseID",arrtestcaselen(tc)&"_Output")
			screenshot=datatable("screenshot",arrtestcaselen(tc)&"_Output")
			Dim arrchildobj
			ReDim arrchildobj(7)
			' jtps
			Set basewin=Browser(strBrwsID).Page(strPageId)
			

			For colcnt = 0 To 6 Step 1
				If colcnt=0 Then
					arrchildobj(0)=strFieldName
				else
				
					if datatable("Field_Name"&colcnt,arrtestcaselen(tc)&"_Output")<>"" then
						arrchildobj(colcnt)=datatable("Field_Name"&colcnt,arrtestcaselen(tc)&"_Output")
					End if
				End if
			Next 
			
		
			
			For 	i = 0 To ubound(arrchildobj) Step 1
				
				If  isempty(arrchildobj(i+1)) =true Then 
					strFieldName=arrchildobj(i)
					Exit for
				End If
							
						
			
				
				If arrchildobj(i) <>"" Then 'loop the objects until the child object is blank/ no value
					
					arrobjtype=split(arrchildobj(i),"_")
					
					
					If environment.value("applicationtype")="Web" Then
						Select Case lcase(arrobjtype(0))
						Case "rb"
							Set basewin=basewin.webRadioButton(arrchildobj(i))
						Case "cb"
							Set basewin=basewin.webcheckbox(arrchildobj(i))
						Case "lst"
							Set basewin=basewin.weblist(arrchildobj(i))
						Case "img"
							Set basewin=basewin.image(arrchildobj(i))
						Case "frm"
							Set basewin=basewin.frame(arrchildobj(i))
						Case "lnk"
							Set basewin=basewin.link(arrchildobj(i))
						Case "btn"
							Set basewin=basewin.webbutton(arrchildobj(i))
						Case "txt"
							Set basewin=basewin.webedit(arrchildobj(i))
						Case "elm"
							Set basewin=basewin.webelement(arrchildobj(i))
						Case "chk"
							Set basewin=basewin.webcheckbox(arrchildobj(i))
						Case "tbl"
							Set basewin=basewin.webtable(arrchildobj(i))
						Case "wnd"
							Set basewin=basewin.webelement(arrchildobj(i))
						End select
					End If
				Else 
					Exit for
				End If				
			Next
			
			
			If strcasedesdummy<>"" Then		 ' To update the excel sheet for each case in the scenario
			If Scenarioname="" Then
				Scenarioname=strcasedesdummy
			End If
			
			If intScriptstartval>1 Then
				'Datatable.ExportSheet environment.value("Results")&Environment.value("Resultsheet")&".xlsx",arrtestcaselen(tc)&"_Output"	
				'Updatesheet arrtestcaselen(tc),Scenarioname,strFaillst,strInstIDreportExcel
				'Updatesheet vartestscenarioname,Scenarioname,strFaillst,strInstIDreportExcel
				strFaillst=""
			End If
			intScriptstartval=intScriptstartval+1
			Scenarioname=strcasedesdummy				
		End If ' closing the test case update in the result excel
		
		If strcasedesdummy<>""  Then
			If strCaseDesc="" Then
				strCaseDesc=strcasedesdummy
			ElseIf strCaseDesc<> strcasedesdummy Then
				strCaseDesc=strcasedesdummy
			End If
			
		End If
		
		
		
		
		
		
		
		set objjtpswin =basewin 
		If boolen="Y" and ucase(boolRun)="Y" and ucase(strAppType)="JTPS"Then			
	'Call Allalertshandlers	
			select case lcase(strAction)
			case "login"
				strAppType=datatable("Application_Type",arrtestcaselen(tc)&"_Output")
				strUrl=datatable("Input",arrtestcaselen(tc)&"_Output")
				strUID=datatable("UID",arrtestcaselen(tc)&"_Output")
				strPWD=datatable("PWD",arrtestcaselen(tc)&"_Output")
				strName=strBrwsID
				strResult=login(strUrl,strAppType,strUID,strPWD,strName)
				strResult1="Log in to application of Portal"
			case "setdata"			
				strResult=jtps_setdata(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Set value into "&strFieldName
			Case "verifydata"
				strResult=jtps_verifydata(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Verify the data in "&strFieldName
			Case "getvalue"
				strResult=jtps_GetValue(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Get value from "&strFieldName
			Case "waitProperty"
				strResult=jtps_WaitProperty(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Wait for object of "&strFieldName
			Case "closewindow"
				strResult=closewindow(strBrwsID)	
				strResult1="Close the browser "&strBrwsID
			case "typedata"			
				strResult=jtps_typedata(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Type the data into "&strFieldName
			Case "objexist"
				strResult=jtps_objExist(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Check the "&strFieldName&" Object is available in "&strPageId
			Case "selectcell"
				strResult=JTPS_Selectcell(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Check the "&strFieldName&" Object is available in "&strPageId
			Case "deleterow"
				strResult=JTPS_DeleteRow(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Check the "&strFieldName&" Object is available in "&strPageId
			Case "addrow"
				strResult=JTPS_AddRow(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Check the "&strFieldName&" Object is available in "&strPageId
			Case "verifycell"
				strResult=JTPS_VerifyCell(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Check the "&strFieldName&" Object is available in "&strPageId	
			Case "wait"
				strResult=fn_wait(strInput)
				StrResult1="sync is successful"
			Case "scrollandclick"	
				strResult=jtps_scrollandclick(strBrwsID,strPageId,strFieldName,strInput)
				strResult1="Scroll Page the "&strFieldName&" Object is available in "&strPageId	
			End select
			
			
		End If
		
			'Write Results
		
		If Ucase(boolen)="Y" and ucase(boolRun)="Y" then
			
			
			arrResult=split(strResult,"\")
			If ucase(strAppType)="JTPS" Then
				strBname=strBrwsID&"\JTPS"		
				
			End If
			
			'Update the Error to Fail
			
			If trim(ucase(arrResult(0)))="ERROR" Then
				arrResult(0)="FAIL"
			End If
			
			'Update the datatable
			'Call ReportLog(arrtestcaselen(tc),strCaseDesc,intstepNum,trim(arrResult(0)),arrResult(1),screenshot)
			Call ReportLog(vartestscenarioname,strCaseDesc,intstepNum,trim(arrResult(0)),arrResult(1),screenshot)
			'Call HTMLReportLog(intstepNum,trim(arrResult(0)),strResult1,arrResult(1),strCaseDesc,arrtestcaselen(tc),screenshot)
			Call HTMLReportLog(intstepNum,trim(arrResult(0)),strResult1,arrResult(1),strCaseDesc,vartestscenarioname,screenshot)
			
			datatable("Message",arrtestcaselen(tc)&"_Output")=arrResult(1)
			datatable("Result",arrtestcaselen(tc)&"_Output")=arrResult(0)
			
			If trim(ucase(arrResult(0)))="FAIL" or trim(ucase(arrResult(0)))="ERROR" Then	
				if strFaillst="" then
					strFaillst=intstepNum
				else
					strFaillst=strFaillst&","&intstepNum
				End if	
			End If
			
			
		End If
	Next
	environment.value("errorexist")="False"
	variteration = variteration+1
	If cint(arrtestcaserep(tc))>0 Then
		If cint(arrtestcaserep(tc))-1<variteration Then
			Exit do
		End If
	ElseIf cint(arrtestcaserep(tc)) =0 Then
		Exit do
	End If
	
	
	set dict=nothing 'undo the dict object for each iteration and reload the config file
	LoadFunctionLibrary environment.value("varpath")&"FunctionLibrary\Config.txt"
	
Loop While  cint(arrtestcaserep(tc))>variteration


Next

Msgbox "Execution Completed; Results can be verified from the Report"



'datatable.Importsheet "F:\UFT_AUTOMATION\Trade360 Automation Scripts\TestData\9.4ATQA\TPS_INWDOC3_Day1.xlsx


'Repositoriescollection.Add "C:\JTPSAutomation\ObjectRepository\JTPS\JTPS.tsr"
'LoadFunctionLibrary "C:\JTPSAutomation\FunctionLibrary\CommonFunctions_JTPS.txt"
''set objjtpswin =Browser("jTPS").Page("pjTPS")
'strBrwsID=""
'strPageId=""
'strFieldName="lnk_MainMenuMore"
'strExpRes=""
'call JTPS_Setdata(strBrwsID,strPageId,strFieldName,strExpRes)



'
'Repositoriescollection.Add "D:\QTPJTPSAutomation\ObjectRepository\JTPS\JTPS.tsr"
'Set a =Browser("jTPS").Page("pjTPS").WebElement("wnd_Reference").Object.getElementsByClassName("ngViewport ng-scope")
'
'


