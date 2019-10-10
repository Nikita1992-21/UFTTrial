
Dim MRowCount , i, ModuleExe,ModuleID,TCRowCount,j,ModuleID2,TestCaseExe,TestCaseID,TCStepRowCount,k,TestCaseID2,keyward

DataTable.AddSheet "Module"
DataTable.AddSheet "TestCase"
DataTable.AddSheet "TestStep"

DataTable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\Hybide_Framework\Organizer\Organizer.xls",1 , 3
DataTable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\Hybide_Framework\Organizer\Organizer.xls",2 , 4
DataTable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\Hybide_Framework\Organizer\Organizer.xls",3 , 5

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
		'MsgBox TestCaseID
		
		TCStepRowCount = DataTable.GetSheet("TestStep").GetRowCount
		For k = 1 To TCStepRowCount Step 1
			DataTable.SetCurrentRow(k)
			TestCaseID2 = DataTable(5,5)
			If TestCaseID =  TestCaseID2 Then
				keyward = DataTable(4,5)
				'MsgBox keyward
				Select Case keyward
					Case "FL"
					DataTable(7,5) = FlightLogin("John","HP")
					Case "CA"
					DataTable(7,5) = Close_App
					Case "FF"
					DataTable(7,5) = FindFlights
					Case "SF"
					DataTable(7,5) = SelectFlight
					Case "CO"
					DataTable(7,5) = CompleteOrder("Suraj")
					
				End Select
			End If
		Next
	End If
Next
	
	
	End If
Next
