'Checking File
Dim objFSO, strFile, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
strFile = objFSO.GetAbsolutePathName(".") & "\data.txt"
If Not objFSO.FileExists(strFile) Then
    objFSO.CreateTextFile(strFile).Close
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /c start /max notepad.exe " & strFile, 1, False
    WScript.sleep 1500
    WScript.Echo "Paste Your Data Here & run the Bot Again"
End If
Set objFSO = Nothing
Set objShell = Nothing
'------------------------------------------------------------------------------'
Dim ExcelFileName, objExcel, objWorkbook, WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
' Specify the Excel file name
ExcelFileName = "data.xlsx"
' Create the Excel.Application object
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
'Skip
On Error GoTo 0
' Check if the workbook is open
On Error Resume Next
Set objWorkbook = objExcel.Workbooks(ExcelFileName)
' Save and close the workbook
If Not objWorkbook Is Nothing Then ' This line checks if the workbook is open
    objWorkbook.Save
    objWorkbook.Close ' This line will only run if the workbook is open
End If
On Error GoTo 0
'----------------------------------------------------------'
'kill xlsx
WshShell.Run "cmd /c taskkill /f /im excel.exe", 0, True
'---------------------------------------------------'
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objFSO = CreateObject("Scripting.FileSystemObject")
' Check if data.xlsx exists
If objFSO.FileExists("data.xlsx") Then
    Set objWorkbook = objExcel.Workbooks.Open(objFSO.GetAbsolutePathName("data.xlsx"))
Else
    Set objWorkbook = objExcel.Workbooks.Add()
End If
Set objWorksheet = objWorkbook.Worksheets(1)
Set objTextFile = objFSO.OpenTextFile("data.txt", 1)
' Find the last row with data in the worksheet
row = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row + 2 ' Start from the row below the last row with data
col = 1
Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    If strNextLine <> "" Then
        objWorksheet.Cells(row, col).Value = strNextLine
        col = col + 1
    End If
    ' If an empty line is encountered, move to the next row but do not append anything
    If strNextLine = "" Then
        If col > 1 Then
            row = row + 1
        End If
        col = 1
    End If
Loop
' Save and close the workbook immediately after appending the data
If objFSO.FileExists(objFSO.GetAbsolutePathName("data.xlsx")) Then 
    objWorkbook.Save 
Else
    objWorkbook.SaveAs objFSO.GetAbsolutePathName("data.xlsx"), 51 ' 51 corresponds to xlsx format 
End If
objWorkbook.Close
objExcel.Quit
'notify 
Set oFSO = CreateObject("Scripting.FileSystemObject")
filePath = "SsS01001pLAY.wav"
If oFSO.FileExists(filePath) Then
    Set oPlayer = CreateObject("WMPlayer.OCX")
    oPlayer.URL = filePath

    Dim fso, tempFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Create a temporary file
    Set tempFile = fso.CreateTextFile("temp.vbs", True)
    ' Write to the file
    tempFile.WriteLine("MsgBox ""COMPLETED"", vbInformation + vbOKOnly, ""Status""")
    ' Close the file
    tempFile.Close
    ' Run the script
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run "temp.vbs", 1, false
    ' Delete the file after a delay
    oPlayer.controls.play
    WScript.Sleep 1000
    oPlayer.controls.stop
    Set oPlayer = Nothing
    fso.DeleteFile("temp.vbs")
End If
Set oFSO = Nothing
'self-destruct
'WshShell.Run "cmd /c taskkill /f /im excel.exe", 0, True
Set objExcel = CreateObject("Excel.Application")
strFileName = "data.xlsx"
strFilePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(strFileName)
objExcel.Visible = True
objExcel.WindowState = -4137 ' -4137 represents the constant for xlMaximized
'objExcel.WindowState = -4140 ' xlMinimized
Set objWorkbook = objExcel.Workbooks.Open(strFilePath)
Set objWorkbook = Nothing
Set objExcel = Nothing
WshShell.Run "cmd /c taskkill /f /im wscript.exe", 0, True
WshShell.Run "cmd /c taskkill /f /im cscript.exe", 0, True