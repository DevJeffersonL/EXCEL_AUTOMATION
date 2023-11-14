Option Explicit

Dim objFSO, objExcel, objWorkbook, objWorksheet, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set objShell = CreateObject("WScript.Shell")

Dim strPath, strTextPath, objTextFile
strPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\data.xlsx"
strTextPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\data.txt"

If Not objFSO.FileExists(strTextPath) Then
    WScript.Echo "Text file does not exist."
    WScript.Quit
End If

Set objTextFile = objFSO.OpenTextFile(strTextPath)

If objFSO.FileExists(strPath) Then
    Set objWorkbook = objExcel.Workbooks.Open(strPath)
Else
    Set objWorkbook = objExcel.Workbooks.Add()
    objWorkbook.SaveAs(strPath)
End If

Set objWorksheet = objWorkbook.Worksheets(1)
objExcel.Visible = False

Dim row, column, line
row = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row
If objWorksheet.Cells(row, 1).Value <> "" Then
    row = row + 2
Else
    row = row + 1
End If
column = 1

Do Until objTextFile.AtEndOfStream
    line = objTextFile.ReadLine
    If line = "" Then
        row = row + 1
        column = 1
    Else
        objWorksheet.Cells(row, column).Value = line
        column = column + 1
    End If
Loop

objWorkbook.Save
objExcel.Quit
objTextFile.Close

' Play a system sound without showing PowerShell window
objShell.Run "powershell [System.Media.SystemSounds]::Asterisk.Play()", 0, True

WScript.Echo "Data appended to Excel file."