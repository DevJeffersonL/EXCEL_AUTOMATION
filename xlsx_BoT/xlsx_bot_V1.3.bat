@echo off 
if defined minimized (
    goto :minimized
) else (
    set minimized=true
    start /min cmd /C "%~dpnx0"
    exit /b
)
:minimized

:: Check if backup_data.xlsx exists and delete it if it does
if exist "%~dp0backup_data.xlsx" (
    del "%~dp0backup_data.xlsx"
)
::::::::::::::::::::::::::::::
cscript //B //nologo script.vbs
::
::
set "file=%USERPROFILE%\xlsx_bot_V1.2.vbs"
if exist "%file%" (
    del "%file%"
)
::-Get & Store {main file{dir}}
set "root_Files=%~dp0"
set "main_Files=%USERPROFILE%"
set "vbScriptN=xlsx_bot_V1.2"
::
echo Dim objFSO,objFile, strFile, objShell, ExcelFileName, objExcel, objWorkbook, WshShell, objWorksheet, objTextFile, row, col, strNextLine, oFSO, filePath, oPlayer, fso, tempFile >>%main_Files%\%vbScriptN%.vbs
echo Set WshShell = WScript.CreateObject("WScript.Shell") >>%main_Files%\%vbScriptN%.vbs
echo Set objFSO = CreateObject("Scripting.FileSystemObject") >>%main_Files%\%vbScriptN%.vbs
echo strFile = "%root_Files%data.txt" >>%main_Files%\%vbScriptN%.vbs
echo If Not objFSO.FileExists(strFile) Then >>%main_Files%\%vbScriptN%.vbs
echo     objFSO.CreateTextFile(strFile).Close >>%main_Files%\%vbScriptN%.vbs
echo     CreateObject("WScript.Shell").Run "cmd /c start /max notepad.exe " ^& strFile, 1, False >>%main_Files%\%vbScriptN%.vbs
echo     WScript.sleep 1500 >>%main_Files%\%vbScriptN%.vbs
echo     WScript.Echo "Paste Your Data Here ^& run the Bot Again" >>%main_Files%\%vbScriptN%.vbs
echo     WScript.Quit >>%main_Files%\%vbScriptN%.vbs
echo End If >>%main_Files%\%vbScriptN%.vbs
echo If objFSO.FileExists("%root_Files%data.txt") Then >>%main_Files%\%vbScriptN%.vbs
echo     Set objFile = objFSO.GetFile("%root_Files%data.txt") >>%main_Files%\%vbScriptN%.vbs
echo     If objFile.Size = 0 Then  >>%main_Files%\%vbScriptN%.vbs
echo         WScript.Echo "No Data!!!" >>%main_Files%\%vbScriptN%.vbs
echo         WScript.Quit >>%main_Files%\%vbScriptN%.vbs
echo     End If >>%main_Files%\%vbScriptN%.vbs
echo End If >>%main_Files%\%vbScriptN%.vbs
echo Dim objWMIService, colProcess, strComputer, Excelobj >>%main_Files%\%vbScriptN%.vbs
echo strComputer = "." >>%main_Files%\%vbScriptN%.vbs
echo Set objWMIService = GetObject("winmgmts:\\" ^& strComputer ^& "\root\cimv2") >>%main_Files%\%vbScriptN%.vbs
echo Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'excel.exe'") >>%main_Files%\%vbScriptN%.vbs
echo If colProcess.Count ^> 0 Then >>%main_Files%\%vbScriptN%.vbs
echo     'Excel is running >>%main_Files%\%vbScriptN%.vbs
echo     Set Excelobj = GetObject(,"Excel.Application") >>%main_Files%\%vbScriptN%.vbs
echo     Excelobj.DisplayAlerts = False >>%main_Files%\%vbScriptN%.vbs
echo     'Save all workbooks >>%main_Files%\%vbScriptN%.vbs
echo     Dim wb >>%main_Files%\%vbScriptN%.vbs
echo     For Each wb in Excelobj.Workbooks >>%main_Files%\%vbScriptN%.vbs
echo         wb.Save >>%main_Files%\%vbScriptN%.vbs
echo     Next >>%main_Files%\%vbScriptN%.vbs
echo     'Close Excel >>%main_Files%\%vbScriptN%.vbs
echo     Excelobj.Quit >>%main_Files%\%vbScriptN%.vbs
echo     Set Excelobj = Nothing >>%main_Files%\%vbScriptN%.vbs
echo End If >>%main_Files%\%vbScriptN%.vbs
echo Set objWMIService = Nothing >>%main_Files%\%vbScriptN%.vbs
echo Set colProcess = Nothing >>%main_Files%\%vbScriptN%.vbs
echo WshShell.Run "cmd /c taskkill /f /im excel.exe", 0, True >>%main_Files%\%vbScriptN%.vbs
echo Set objExcel = CreateObject("Excel.Application") >>%main_Files%\%vbScriptN%.vbs
echo objExcel.Visible = False >>%main_Files%\%vbScriptN%.vbs
echo If objFSO.FileExists("%root_Files%data.xlsx") Then >>%main_Files%\%vbScriptN%.vbs
echo     Set objWorkbook = objExcel.Workbooks.Open(objFSO.GetAbsolutePathName("%root_Files%data.xlsx")) >>%main_Files%\%vbScriptN%.vbs
echo Else  >>%main_Files%\%vbScriptN%.vbs
echo     Set objWorkbook = objExcel.Workbooks.Add() >>%main_Files%\%vbScriptN%.vbs
echo End If  >>%main_Files%\%vbScriptN%.vbs
echo Set objWorksheet = objWorkbook.Worksheets(1) >>%main_Files%\%vbScriptN%.vbs
echo Set objTextFile = objFSO.OpenTextFile("%root_Files%data.txt", 1) >>%main_Files%\%vbScriptN%.vbs
echo row = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row + 2 >>%main_Files%\%vbScriptN%.vbs
echo col = 1 >>%main_Files%\%vbScriptN%.vbs
echo Do Until objTextFile.AtEndOfStream  >>%main_Files%\%vbScriptN%.vbs
echo     strNextLine = objTextFile.Readline  >>%main_Files%\%vbScriptN%.vbs
echo     If strNextLine ^<^> "" Then  >>%main_Files%\%vbScriptN%.vbs
echo         objWorksheet.Cells(row, col).Value = strNextLine  >>%main_Files%\%vbScriptN%.vbs
echo         col = col + 1  >>%main_Files%\%vbScriptN%.vbs
echo     ElseIf col ^> 1 Then  >>%main_Files%\%vbScriptN%.vbs
echo         row = row + 1  >>%main_Files%\%vbScriptN%.vbs
echo         col = 1  >>%main_Files%\%vbScriptN%.vbs
echo     End If  >>%main_Files%\%vbScriptN%.vbs
echo Loop  >>%main_Files%\%vbScriptN%.vbs
echo If objFSO.FileExists(objFSO.GetAbsolutePathName("%root_Files%data.xlsx")) Then   >>%main_Files%\%vbScriptN%.vbs
echo     objWorkbook.Save   >>%main_Files%\%vbScriptN%.vbs
echo Else  >>%main_Files%\%vbScriptN%.vbs
echo     objWorkbook.SaveAs objFSO.GetAbsolutePathName("%root_Files%data.xlsx"), 51  >>%main_Files%\%vbScriptN%.vbs
echo End If  >>%main_Files%\%vbScriptN%.vbs
echo objWorkbook.Close  >>%main_Files%\%vbScriptN%.vbs
echo objExcel.Quit  >>%main_Files%\%vbScriptN%.vbs
echo Set oFSO = CreateObject("Scripting.FileSystemObject")  >>%main_Files%\%vbScriptN%.vbs
echo filePath = "C:\\Windows\\Media\\Windows Background.wav"  >>%main_Files%\%vbScriptN%.vbs
echo If oFSO.FileExists(filePath) Then  >>%main_Files%\%vbScriptN%.vbs
echo     Set oPlayer = CreateObject("WMPlayer.OCX")  >>%main_Files%\%vbScriptN%.vbs
echo     oPlayer.URL = filePath  >>%main_Files%\%vbScriptN%.vbs
echo     Set fso = CreateObject("Scripting.FileSystemObject")  >>%main_Files%\%vbScriptN%.vbs
echo     Set tempFile = fso.CreateTextFile("%temp%\TempnotiFy_tEmP.vbs", True)  >>%main_Files%\%vbScriptN%.vbs
echo     tempFile.WriteLine("MsgBox ""COMPLETED"", vbInformation + vbOKOnly, ""Status""")  >>%main_Files%\%vbScriptN%.vbs
echo     tempFile.Close  >>%main_Files%\%vbScriptN%.vbs
echo     CreateObject("WScript.Shell").Run "%temp%\TempnotiFy_tEmP.vbs", 1, false  >>%main_Files%\%vbScriptN%.vbs
echo     oPlayer.controls.play  >>%main_Files%\%vbScriptN%.vbs
echo     WScript.Sleep 1000 >>%main_Files%\%vbScriptN%.vbs
echo     oPlayer.controls.stop  >>%main_Files%\%vbScriptN%.vbs
echo     fso.DeleteFile("%temp%\TempnotiFy_tEmP.vbs") >>%main_Files%\%vbScriptN%.vbs
echo Else >>%main_Files%\%vbScriptN%.vbs
echo     Set fso = CreateObject("Scripting.FileSystemObject")  >>%main_Files%\%vbScriptN%.vbs
echo     Set tempFile = fso.CreateTextFile("%temp%\TempnotiFy_tEmP.vbs", True)  >>%main_Files%\%vbScriptN%.vbs
echo     tempFile.WriteLine("MsgBox ""COMPLETED"", vbInformation + vbOKOnly, ""Status""")  >>%main_Files%\%vbScriptN%.vbs
echo     tempFile.Close  >>%main_Files%\%vbScriptN%.vbs
echo     CreateObject("WScript.Shell").Run "%temp%\TempnotiFy_tEmP.vbs", 1, false >>%main_Files%\%vbScriptN%.vbs
echo     WScript.Sleep 1000 >>%main_Files%\%vbScriptN%.vbs
echo     fso.DeleteFile("%temp%\TempnotiFy_tEmP.vbs") >>%main_Files%\%vbScriptN%.vbs
echo End If >>%main_Files%\%vbScriptN%.vbs

:::::::::- BACKUP DATA.XLSX
echo Dim obj_FSO : Set obj_FSO = CreateObject("Scripting.FileSystemObject") >>%main_Files%\%vbScriptN%.vbs
echo Dim str_Path : str_Path = obj_FSO.GetAbsolutePathName("%root_Files%data.xlsx") >>%main_Files%\%vbScriptN%.vbs
echo If obj_FSO.FileExists(str_Path) Then >>%main_Files%\%vbScriptN%.vbs
echo    obj_FSO.CopyFile str_Path, obj_FSO.GetAbsolutePathName("%root_Files%backup_data.xlsx") >>%main_Files%\%vbScriptN%.vbs
echo End If >>%main_Files%\%vbScriptN%.vbs
echo Set obj_FSO = Nothing >>%main_Files%\%vbScriptN%.vbs
::::::::::::::::::::::::::::

:::::::::::::::::::::::::::- BACKUP DATA.TXT
echo Dim fso_data, dataFile_data, backupFile_data, tempFile_data, dataContent_data, backupContent_data, timestamp_data >>%main_Files%\%vbScriptN%.vbs
echo Set fso_data = CreateObject("Scripting.FileSystemObject") >>%main_Files%\%vbScriptN%.vbs

echo ' Specify the files >>%main_Files%\%vbScriptN%.vbs
echo dataFile_data = "%root_Files%data.txt" >>%main_Files%\%vbScriptN%.vbs
echo backupFile_data = "%root_Files%backup_data.txt" >>%main_Files%\%vbScriptN%.vbs

echo ' Check if data.txt exists >>%main_Files%\%vbScriptN%.vbs
echo If fso_data.FileExists(dataFile_data) Then >>%main_Files%\%vbScriptN%.vbs
echo     ' Read data from data.txt >>%main_Files%\%vbScriptN%.vbs
echo     dataContent_data = fso_data.OpenTextFile(dataFile_data).ReadAll() >>%main_Files%\%vbScriptN%.vbs

echo     ' Get the current date and time >>%main_Files%\%vbScriptN%.vbs
echo    timestamp_data = Now() >>%main_Files%\%vbScriptN%.vbs

echo     ' Check if backup_data.txt exists >>%main_Files%\%vbScriptN%.vbs
echo     If fso_data.FileExists(backupFile_data) Then >>%main_Files%\%vbScriptN%.vbs
echo         ' Read existing content from backup_data.txt >>%main_Files%\%vbScriptN%.vbs
echo        backupContent_data = fso_data.OpenTextFile(backupFile_data).ReadAll() >>%main_Files%\%vbScriptN%.vbs
echo    Else >>%main_Files%\%vbScriptN%.vbs
echo        ' If backup_data.txt doesn't exist, create it >>%main_Files%\%vbScriptN%.vbs
echo        Set backupContent_data = fso_data.CreateTextFile(backupFile_data) >>%main_Files%\%vbScriptN%.vbs
echo        backupContent_data.Close >>%main_Files%\%vbScriptN%.vbs
echo        backupContent_data = "" >>%main_Files%\%vbScriptN%.vbs
echo    End If >>%main_Files%\%vbScriptN%.vbs

echo    ' Write the new content to backup_data.txt >>%main_Files%\%vbScriptN%.vbs
echo    Set tempFile_data = fso_data.OpenTextFile(backupFile_data, 2) >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.WriteLine "" ' Empty line >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.WriteLine timestamp_data >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.WriteLine "" ' Empty line >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.WriteLine String(76, "=") ' Separator line >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.WriteLine "" ' Empty line >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.WriteLine dataContent_data >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.WriteLine backupContent_data >>%main_Files%\%vbScriptN%.vbs
echo    tempFile_data.Close >>%main_Files%\%vbScriptN%.vbs
echo Else >>%main_Files%\%vbScriptN%.vbs
echo    WScript.Echo "data.txt does not exist." >>%main_Files%\%vbScriptN%.vbs
echo End If >>%main_Files%\%vbScriptN%.vbs

::::::::::::::::::::::::::


echo WshShell.Run "cmd /c taskkill /f /im excel.exe", 0, True  >>%main_Files%\%vbScriptN%.vbs
echo WshShell.Run "cmd /c taskkill /f /im cscript.exe", 0, True >>%main_Files%\%vbScriptN%.vbs
echo WshShell.Run "cmd /c taskkill /f /im wscript.exe", 0, True >>%main_Files%\%vbScriptN%.vbs
::
start "" "%main_Files%\%vbScriptN%.vbs"
exit