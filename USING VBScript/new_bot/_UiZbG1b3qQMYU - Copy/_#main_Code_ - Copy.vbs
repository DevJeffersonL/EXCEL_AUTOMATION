Dim objFSO,objFile, strFile, objShell, ExcelFileName, objExcel, objWorkbook, WshShell, objWorksheet, objTextFile, row, col, strNextLine, oFSO, filePath, oPlayer, fso, tempFile 
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject") 
strFile = "data.txt" 
If Not objFSO.FileExists(strFile) Then 
    objFSO.CreateTextFile(strFile).Close 
    CreateObject("WScript.Shell").Run "cmd /c start /max notepad.exe " & strFile, 1, False 
    WScript.sleep 1500 
    WScript.Echo "Paste Your Data Here & run the Bot Again" 
    WScript.Quit 
End If 
If objFSO.FileExists("data.txt") Then 
    Set objFile = objFSO.GetFile("data.txt") 
    If objFile.Size = 0 Then  
        WScript.Echo "No Data!!!" 
        WScript.Quit 
    End If 
End If 

Dim objWMIService, colProcess, strComputer, Excelobj
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'excel.exe'")
If colProcess.Count > 0 Then
    'Excel is running
    Set Excelobj = GetObject(,"Excel.Application")
    Excelobj.DisplayAlerts = False
    'Save all workbooks
    Dim wb
    For Each wb in Excelobj.Workbooks
        wb.Save
    Next
    'Close Excel
    Excelobj.Quit
    Set Excelobj = Nothing
End If
Set objWMIService = Nothing
Set colProcess = Nothing


WshShell.Run "cmd /c taskkill /f /im excel.exe", 0, True 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = False 
If objFSO.FileExists("data.xlsx") Then 
    Set objWorkbook = objExcel.Workbooks.Open(objFSO.GetAbsolutePathName("data.xlsx")) 
Else 
    Set objWorkbook = objExcel.Workbooks.Add() 
End If 
Set objWorksheet = objWorkbook.Worksheets(1) 
Set objTextFile = objFSO.OpenTextFile("data.txt", 1) 
row = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row + 2 
col = 1 
Do Until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.Readline 
    If strNextLine <> "" Then 
        objWorksheet.Cells(row, col).Value = strNextLine 
        col = col + 1 
    ElseIf col > 1 Then 
        row = row + 1 
        col = 1 
    End If 
Loop 
If objFSO.FileExists(objFSO.GetAbsolutePathName("data.xlsx")) Then  
    objWorkbook.Save  
Else 
    objWorkbook.SaveAs objFSO.GetAbsolutePathName("data.xlsx"), 51 
End If 
objWorkbook.Close 
objExcel.Quit 
Set oFSO = CreateObject("Scripting.FileSystemObject") 
filePath = "SsS01001pLAY.wav" 
If oFSO.FileExists(filePath) Then 
    Set oPlayer = CreateObject("WMPlayer.OCX") 
    oPlayer.URL = filePath 
    Set fso = CreateObject("Scripting.FileSystemObject") 
    Set tempFile = fso.CreateTextFile("TempnotiFy_tEmP.vbs", True) 
    tempFile.WriteLine("MsgBox ""COMPLETED"", vbInformation + vbOKOnly, ""Status""") 
    tempFile.Close 
    CreateObject("WScript.Shell").Run "TempnotiFy_tEmP.vbs", 1, false 
    oPlayer.controls.play 
    WScript.Sleep 1000
    oPlayer.controls.stop 
    fso.DeleteFile("TempnotiFy_tEmP.vbs") 
End If 
WshShell.Run "cmd /c taskkill /f /im wscript.exe", 0, True 
WshShell.Run "cmd /c taskkill /f /im cscript.exe", 0, True 
