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