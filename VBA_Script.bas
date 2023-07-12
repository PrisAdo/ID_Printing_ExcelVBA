Attribute VB_Name = "Module2"

Sub Test1()
   Set obj = New Selenium.ChromeDriver
   
   obj.AddArgument "--kiosk-printing"
   
   obj.Start
   obj.Get "https://arslanh.neocities.org/idcard/"
   
   Application.Wait (Now + TimeValue("0:00:30"))
   'Set fields and clear everything
   
   For intRow = 71 To 71
        obj.FindElementById("imgInp").SendKeys (ThisWorkbook.Sheets("Sheet1").Range("A" & intRow).Value)
        obj.FindElementById("firstNameInput").SendKeys (ThisWorkbook.Sheets("Sheet1").Range("B" & intRow).Value)
        obj.FindElementById("lastNameInput").SendKeys (ThisWorkbook.Sheets("Sheet1").Range("C" & intRow).Value)
        obj.FindElementById("positionInput").SendKeys (ThisWorkbook.Sheets("Sheet1").Range("D" & intRow).Value)
        obj.FindElementById("idNumberInput").SendKeys (ThisWorkbook.Sheets("Sheet1").Range("E" & intRow).Value)
    
   
        SendKeys "^p"
        Application.Wait (Now + TimeValue("0:00:05"))
        
         obj.FindElementById("firstNameInput").Clear
        obj.FindElementById("lastNameInput").Clear
        obj.FindElementById("positionInput").Clear
        obj.FindElementById("idNumberInput").Clear
   
   Next
   
End Sub

