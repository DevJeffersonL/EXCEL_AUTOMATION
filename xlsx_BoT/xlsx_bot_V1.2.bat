@echo off 
:minimized
if not "%minimized%"=="" goto :minimized
set minimized=true
start /min cmd /C "%~dpnx0"
goto :EOF
:minimized
::
::-> If exist(Open)
set "file=%USERPROFILE%\xlsx_bot_V1.2.vbs"
if exist "%file%" (
    start "" "%file%"
    exit
)
::-Get & Store {main file{dir}}
set "root_Files=%~dp0"
set "main_Files=%USERPROFILE%"
set "vbScriptN=xlsx_bot_V1.2"
::set "script_name=%~nx0"
::if exist %USERPROFILE%\%vbScriptN%.vbs del /Q "%USERPROFILE%\%vbScriptN%.vbs"
::if exist "%main_Files%" rmdir /S /Q "%main_Files%"
::if not exist "%main_Files%" mkdir "%main_Files%"
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
echo WshShell.Run "cmd /c taskkill /f /im excel.exe", 0, True  >>%main_Files%\%vbScriptN%.vbs
echo WshShell.Run "cmd /c taskkill /f /im cscript.exe", 0, True >>%main_Files%\%vbScriptN%.vbs
echo WshShell.Run "cmd /c taskkill /f /im wscript.exe", 0, True >>%main_Files%\%vbScriptN%.vbs
::
start "" "%main_Files%\%vbScriptN%.vbs"
exit