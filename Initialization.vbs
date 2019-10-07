Dim objUFT
Set objUFT = CreateObject("QuickTest.Application")

objUFT.Visible = True

objUFT.Launch
objUFT.Open "C:\Users\Nikita.Chaudhari\Desktop\Hybide_Framework\Drivers\Driver1"
objUFT.Test.Run
objUFT.Test.Close

objUFT.Quit
Set objUFT = Nothing

