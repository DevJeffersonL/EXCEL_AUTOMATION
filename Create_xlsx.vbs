' Create an Excel application object
Set objExcel = CreateObject("Excel.Application")

' Create a new workbook
Set objWorkbook = objExcel.Workbooks.Add()

' Get the path of the script
strScriptPath = WScript.ScriptFullName
strScriptDirectory = Left(strScriptPath, InStrRev(strScriptPath, "\"))

' Build the full path for saving the workbook
strSavePath = strScriptDirectory & "Data.xlsx"

' Save the workbook to the script's directory
objWorkbook.SaveAs strSavePath

' Close the workbook
objWorkbook.Close

' Quit Excel
objExcel.Quit

' Release the object references
Set objWorkbook = Nothing
Set objExcel = Nothing

' Play a pop-up sound
Set objShell = CreateObject("WScript.Shell")
objShell.Run "mshta ""javascript:var sh=new ActiveXObject('WScript.Shell');sh.Popup('Completed!', 5, 'Success', 64);close();""", 1, True
Set objShell = Nothing