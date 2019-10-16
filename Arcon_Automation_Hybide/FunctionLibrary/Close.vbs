
'**************************************************
'Function Name : Login
'Author : Nikita
'Date of Creation : 
'Date of Modification:
'Input : Username , Password
'Output : My Service Page
'Purpose : Login 
'**************************************************



Function Close_Arcon()
If Browser("Privileged Account Management").Page("Privileged Account Management").Exist(2) Then
Browser("Privileged Account Management").Page("Privileged Account Management").Image("user1").Check CheckPoint("user1")
Browser("Privileged Account Management").Page("Privileged Account Management").Link("Milind").Click
Browser("Privileged Account Management").Page("Privileged Account Management").Link("Logout").Click
Close_Arcon = "Logout Successfully-Passed"
MsgBox Close_Arcon
Else
Close_Arcon = "Logout Unsuccessful-Failed"
End  If
End Function


