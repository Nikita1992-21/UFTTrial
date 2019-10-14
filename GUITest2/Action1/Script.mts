Dim LRowCount, i


datatable.AddSheet "Login"

datatable.ImportSheet "C:\Users\Nikita.Chaudhari\Desktop\TestData\AconLogin_2_.xls",1 ,3

LRowCount = datatable.GetSheet("Login").GetRowCount
 Browser("This site isn’t secure").Page("This site isn’t secure").Link("More information").Click
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
Browser("This site isn’t secure").Back

 	
ElseIf  Browser("Privileged Account Management_3").Page("Privileged Account Management").WebButton("×").Exist(2) Then
'hWnd1 = Browser("Privileged Account Management_3").Page("Privileged Account Management").WebElement("Invalid User Name OR Password").GetROProperty("hWnd") @@ hightlight id_;_Browser("Privileged Account Management 3").Page("Privileged Account Management").WebElement("Invalid User Name OR Password")_;_script infofile_;_ZIP::ssf13.xml_;_
 @@ hightlight id_;_Browser("Website restore error").Page("Website restore error").Image("Info icon")_;_script infofile_;_ZIP::ssf15.xml_;_
Browser("Privileged Account Management_3").Page("Privileged Account Management").WebButton("×").Click @@ hightlight id_;_Browser("Privileged Account Management 3").Page("Privileged Account Management").WebButton("×")_;_script infofile_;_ZIP::ssf29.xml_;_

Browser("This site isn’t secure").Back @@ hightlight id_;_134662_;_script infofile_;_ZIP::ssf39.xml_;_
Browser("This site isn’t secure").Page("Privileged Account Management").Sync @@ hightlight id_;_Browser("This site isn’t secure").Page("Privileged Account Management")_;_script infofile_;_ZIP::ssf40.xml_;_
Browser("This site isn’t secure").Back
'Browser("Privileged Account Management_3").CloseAllTabs
'SystemUtil.CloseProcessByHwnd (hWnd1)
ElseIf Browser("Website restore error").Page("Website restore error").Image("Info icon").Exist(2) Then

Browser("Privileged Account Management_3").Page("This site isn’t secure").Sync @@ hightlight id_;_Browser("Privileged Account Management_3").Page("This site isn’t secure")_;_script infofile_;_ZIP::ssf19.xml_;_
Browser("Privileged Account Management_3").CloseAllTabs @@ hightlight id_;_329950_;_script infofile_;_ZIP::ssf20.xml_;_
Window("Internet Explorer").Close
Else
Browser("Website restore error").Dialog("Message from webpage").WinButton("OK").Click @@ hightlight id_;_1115396_;_script infofile_;_ZIP::ssf44.xml_;_

'MsgBox "Login unsuccessful"
End If

 Next

 @@ hightlight id_;_Browser("This site isn’t secure").Page("Privileged Account Management").WebButton("×")_;_script infofile_;_ZIP::ssf38.xml_;_
 @@ hightlight id_;_134662_;_script infofile_;_ZIP::ssf41.xml_;_
 @@ hightlight id_;_Browser("This site isn’t secure").Page("Privileged Account Management 2").Link("Logout")_;_script infofile_;_ZIP::ssf7.xml_;_
 @@ hightlight id_;_65774_;_script infofile_;_ZIP::ssf15.xml_;_
 @@ hightlight id_;_65774_;_script infofile_;_ZIP::ssf23.xml_;_
