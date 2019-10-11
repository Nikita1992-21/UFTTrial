Dim LRowCount, i


datatable.AddSheet "Login"

datatable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\TestData\AconLogin_1_.xls",1 ,3

LRowCount = datatable.GetSheet("Login").GetRowCount
Browser("This site isn’t secure").Page("This site isn’t secure").Link("More information").Click @@ hightlight id_;_Browser("This site isn’t secure").Page("This site isn’t secure").Link("More information")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("This site isn’t secure").Page("This site isn’t secure").Link("Go on to the webpage (not").Click

 For i = 1 To LRowCount Step 1
 
datatable.SetCurrentRow(i)
     Browser("This site isn’t secure").Page("Privileged Account Management").WebEdit("ctl00$MainContent$txtusername").Set datatable(1,3)
    Browser("This site isn’t secure").Page("Privileged Account Management").WebEdit("ctl00$MainContent$txtpassword").SetSecure datatable(2,3) @@ hightlight id_;_Browser("This site isn’t secure").Page("Privileged Account Management").WebEdit("ctl00$MainContent$txtpassword")_;_script infofile_;_ZIP::ssf4.xml_;_
    Browser("This site isn’t secure").Page("Privileged Account Management").Image("ctl00$MainContent$imgbtnLogin").Click 24,12 @@ hightlight id_;_Browser("This site isn’t secure").Page("Privileged Account Management").Image("ctl00$MainContent$imgbtnLogin")_;_script infofile_;_ZIP::ssf5.xml_;_
    
If Browser("This site isn’t secure").Page("Privileged Account Management_2").WebElement("ctl00_MainContent_lblHeader").Exist(2) Then
'Msgbox "Login Successful"
Browser("This site isn’t secure").Page("Privileged Account Management_2").Link("Milind").Click
Browser("This site isn’t secure").Page("Privileged Account Management_2").Link("Logout").Click
 	
Else 


hWnd = Browser("Privileged Account Management_2").Dialog("Message from webpage").GetROProperty("hWnd")
SystemUtil.CloseProcessByHwnd (hWnd)
Browser("This site isn’t secure").Page("This site isn’t secure").Link("More information").Click @@ hightlight id_;_Browser("This site isn’t secure").Page("This site isn’t secure").Link("More information")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("This site isn’t secure").Page("This site isn’t secure").Link("Go on to the webpage (not").Click
'systemutil.CloseProcessByHwnd (hwnd)
'MsgBox "Login unsuccessful"
End If

 Next

 @@ hightlight id_;_Browser("This site isn’t secure").Page("Privileged Account Management 2").Link("Logout")_;_script infofile_;_ZIP::ssf7.xml_;_
