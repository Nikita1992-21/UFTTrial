
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
Browser("This site isn’t secure").Page("Privileged Account Management_2").Link("Milind").Click
Browser("This site isn’t secure").Page("Privileged Account Management_2").Link("Logout").Click	
Close_Arcon = "Logout Successfully-Passed"
End Function
