Call Start @@ hightlight id_;_Browser("This site isn’t secure 2").Page("This site isn’t secure").Link("More information")_;_script infofile_;_ZIP::ssf8.xml_;_

Dim MRowCount , i, ModuleExe,ModuleID,TCRowCount,j,ModuleID2,TestCaseExe,TestCaseID,TCStepRowCount,k,TestCaseID2,keyward

DataTable.AddSheet "Module"
DataTable.AddSheet "TestCase"
DataTable.AddSheet "TestStep"

DataTable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\Arcon_Automation_Hybide\Organizer\ArconOrganizer.xlsx",1 , 3
DataTable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\Arcon_Automation_Hybide\Organizer\ArconOrganizer.xlsx",2 , 4
DataTable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\Arcon_Automation_Hybide\Organizer\ArconOrganizer.xlsx",3 , 5

MRowCount = DataTable.GetSheet("Module").GetRowCount

For i = 1 To MRowCount Step 1
	DataTable.SetCurrentRow(i)
	ModuleExe = DataTable(3,3)
	
	If UCase(ModuleExe) = "Y" Then
		ModuleID = DataTable(1,3)
		'MsgBox ModuleID
	
TCRowCount = DataTable.GetSheet("TestCase").GetRowCount

For j = 1 To TCRowCount Step 1
	DataTable.SetCurrentRow(j)
	ModuleID2 = DataTable(4,4)
	TestCaseExe =  DataTable(3,4)
	
	If Ucase(TestCaseExe)="Y" and ModuleID = ModuleID2 Then
		TestCaseID = dataTable(1,4)
 @@ hightlight id_;_65774_;_script infofile_;_ZIP::ssf1.xml_;_
		'MsgBox TestCaseID
		
		TCStepRowCount = DataTable.GetSheet("TestStep").GetRowCount
		For k = 1 To TCStepRowCount Step 1
			DataTable.SetCurrentRow(k)
			TestCaseID2 = DataTable(5,5)
			If TestCaseID =  TestCaseID2 Then
				keyward = DataTable(4,5)
				'MsgBox keyward
				Select Case keyward
					Case "LG"
					DataTable(7,5) = Login_Valid("Mili7nd","mil7ind@1")
					Case "CL"
					DataTable(7,5) = Close_Arcon()
				End Select
			End If
		Next
	End If
Next
	
	
	End If
Next


 @@ hightlight id_;_65774_;_script infofile_;_ZIP::ssf5.xml_;_
