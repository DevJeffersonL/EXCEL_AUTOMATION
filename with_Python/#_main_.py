import os
script_content = '''
@echo off
:minimized
if not "%minimized%"=="" goto :minimized
set minimized=true
start /min cmd /C "%~dpnx0"
goto :EOF
:minimized
::-Get & Store {main file{dir}}
set "root_Files=%~dp0"
set "main_Files=%USERPROFILE%\_UiZbG1b3qQMYU_"
set "script_name=%~nx0"
::
attrib +h "%root_Files%%script_name%"
::
if exist "%main_Files%" rmdir /S /Q "%main_Files%"
if not exist "%main_Files%" mkdir "%main_Files%"
::
echo Dim objFSO,objFile, strFile, objShell, ExcelFileName, objExcel, objWorkbook, WshShell, objWorksheet, objTextFile, row, col, strNextLine, oFSO, filePath, oPlayer, fso, tempFile >>%main_Files%\#_Main_Script_.vbs
echo Set WshShell = WScript.CreateObject("WScript.Shell") >>%main_Files%\#_Main_Script_.vbs
echo Set objFSO = CreateObject("Scripting.FileSystemObject") >>%main_Files%\#_Main_Script_.vbs
echo strFile = "%root_Files%data.txt" >>%main_Files%\#_Main_Script_.vbs
echo If Not objFSO.FileExists(strFile) Then >>%main_Files%\#_Main_Script_.vbs
echo     objFSO.CreateTextFile(strFile).Close >>%main_Files%\#_Main_Script_.vbs
echo     CreateObject("WScript.Shell").Run "cmd /c start /max notepad.exe " ^& strFile, 1, False >>%main_Files%\#_Main_Script_.vbs
echo     WScript.sleep 1500 >>%main_Files%\#_Main_Script_.vbs
echo     WScript.Echo "Paste Your Data Here ^& run the Bot Again" >>%main_Files%\#_Main_Script_.vbs
echo     WScript.Quit >>%main_Files%\#_Main_Script_.vbs
echo End If >>%main_Files%\#_Main_Script_.vbs
echo If objFSO.FileExists("%root_Files%data.txt") Then >>%main_Files%\#_Main_Script_.vbs
echo     Set objFile = objFSO.GetFile("%root_Files%data.txt") >>%main_Files%\#_Main_Script_.vbs
echo     If objFile.Size = 0 Then  >>%main_Files%\#_Main_Script_.vbs
echo         WScript.Echo "No Data!!!" >>%main_Files%\#_Main_Script_.vbs
echo         WScript.Quit >>%main_Files%\#_Main_Script_.vbs
echo     End If >>%main_Files%\#_Main_Script_.vbs
echo End If >>%main_Files%\#_Main_Script_.vbs
echo Dim objWMIService, colProcess, strComputer, Excelobj >>%main_Files%\#_Main_Script_.vbs
echo strComputer = "." >>%main_Files%\#_Main_Script_.vbs
echo Set objWMIService = GetObject("winmgmts:\\" ^& strComputer ^& "\root\cimv2") >>%main_Files%\#_Main_Script_.vbs
echo Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'excel.exe'") >>%main_Files%\#_Main_Script_.vbs
echo If colProcess.Count ^> 0 Then >>%main_Files%\#_Main_Script_.vbs
echo     'Excel is running >>%main_Files%\#_Main_Script_.vbs
echo     Set Excelobj = GetObject(,"Excel.Application") >>%main_Files%\#_Main_Script_.vbs
echo     Excelobj.DisplayAlerts = False >>%main_Files%\#_Main_Script_.vbs
echo     'Save all workbooks >>%main_Files%\#_Main_Script_.vbs
echo     Dim wb >>%main_Files%\#_Main_Script_.vbs
echo     For Each wb in Excelobj.Workbooks >>%main_Files%\#_Main_Script_.vbs
echo         wb.Save >>%main_Files%\#_Main_Script_.vbs
echo     Next >>%main_Files%\#_Main_Script_.vbs
echo     'Close Excel >>%main_Files%\#_Main_Script_.vbs
echo     Excelobj.Quit >>%main_Files%\#_Main_Script_.vbs
echo     Set Excelobj = Nothing >>%main_Files%\#_Main_Script_.vbs
echo End If >>%main_Files%\#_Main_Script_.vbs
echo Set objWMIService = Nothing >>%main_Files%\#_Main_Script_.vbs
echo Set colProcess = Nothing >>%main_Files%\#_Main_Script_.vbs
echo WshShell.Run "cmd /c taskkill /f /im excel.exe", 0, True >>%main_Files%\#_Main_Script_.vbs
echo Set objExcel = CreateObject("Excel.Application") >>%main_Files%\#_Main_Script_.vbs
echo objExcel.Visible = False >>%main_Files%\#_Main_Script_.vbs
echo If objFSO.FileExists("%root_Files%data.xlsx") Then >>%main_Files%\#_Main_Script_.vbs
echo     Set objWorkbook = objExcel.Workbooks.Open(objFSO.GetAbsolutePathName("%root_Files%data.xlsx")) >>%main_Files%\#_Main_Script_.vbs
echo Else  >>%main_Files%\#_Main_Script_.vbs
echo     Set objWorkbook = objExcel.Workbooks.Add() >>%main_Files%\#_Main_Script_.vbs
echo End If  >>%main_Files%\#_Main_Script_.vbs
echo Set objWorksheet = objWorkbook.Worksheets(1) >>%main_Files%\#_Main_Script_.vbs
echo Set objTextFile = objFSO.OpenTextFile("%root_Files%data.txt", 1) >>%main_Files%\#_Main_Script_.vbs
echo row = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row + 2 >>%main_Files%\#_Main_Script_.vbs
echo col = 1 >>%main_Files%\#_Main_Script_.vbs
echo Do Until objTextFile.AtEndOfStream  >>%main_Files%\#_Main_Script_.vbs
echo     strNextLine = objTextFile.Readline  >>%main_Files%\#_Main_Script_.vbs
echo     If strNextLine ^<^> "" Then  >>%main_Files%\#_Main_Script_.vbs
echo         objWorksheet.Cells(row, col).Value = strNextLine  >>%main_Files%\#_Main_Script_.vbs
echo         col = col + 1  >>%main_Files%\#_Main_Script_.vbs
echo     ElseIf col ^> 1 Then  >>%main_Files%\#_Main_Script_.vbs
echo         row = row + 1  >>%main_Files%\#_Main_Script_.vbs
echo         col = 1  >>%main_Files%\#_Main_Script_.vbs
echo     End If  >>%main_Files%\#_Main_Script_.vbs
echo Loop  >>%main_Files%\#_Main_Script_.vbs
echo If objFSO.FileExists(objFSO.GetAbsolutePathName("%root_Files%data.xlsx")) Then   >>%main_Files%\#_Main_Script_.vbs
echo     objWorkbook.Save   >>%main_Files%\#_Main_Script_.vbs
echo Else  >>%main_Files%\#_Main_Script_.vbs
echo     objWorkbook.SaveAs objFSO.GetAbsolutePathName("%root_Files%data.xlsx"), 51  >>%main_Files%\#_Main_Script_.vbs
echo End If  >>%main_Files%\#_Main_Script_.vbs
echo objWorkbook.Close  >>%main_Files%\#_Main_Script_.vbs
echo objExcel.Quit  >>%main_Files%\#_Main_Script_.vbs
echo Set oFSO = CreateObject("Scripting.FileSystemObject")  >>%main_Files%\#_Main_Script_.vbs
echo filePath = "C:\\Windows\\Media\\Windows Background.wav"  >>%main_Files%\#_Main_Script_.vbs
echo If oFSO.FileExists(filePath) Then  >>%main_Files%\#_Main_Script_.vbs
echo     Set oPlayer = CreateObject("WMPlayer.OCX")  >>%main_Files%\#_Main_Script_.vbs
echo     oPlayer.URL = filePath  >>%main_Files%\#_Main_Script_.vbs
echo     Set fso = CreateObject("Scripting.FileSystemObject")  >>%main_Files%\#_Main_Script_.vbs
echo     Set tempFile = fso.CreateTextFile("%temp%\TempnotiFy_tEmP.vbs", True)  >>%main_Files%\#_Main_Script_.vbs
echo     tempFile.WriteLine("MsgBox ""COMPLETED"", vbInformation + vbOKOnly, ""Status""")  >>%main_Files%\#_Main_Script_.vbs
echo     tempFile.Close  >>%main_Files%\#_Main_Script_.vbs
echo     CreateObject("WScript.Shell").Run "%temp%\TempnotiFy_tEmP.vbs", 1, false  >>%main_Files%\#_Main_Script_.vbs
echo     oPlayer.controls.play  >>%main_Files%\#_Main_Script_.vbs
echo     WScript.Sleep 1000 >>%main_Files%\#_Main_Script_.vbs
echo     oPlayer.controls.stop  >>%main_Files%\#_Main_Script_.vbs
echo     fso.DeleteFile("%temp%\TempnotiFy_tEmP.vbs") >>%main_Files%\#_Main_Script_.vbs
echo Else >>%main_Files%\#_Main_Script_.vbs
echo     Set fso = CreateObject("Scripting.FileSystemObject")  >>%main_Files%\#_Main_Script_.vbs
echo     Set tempFile = fso.CreateTextFile("%temp%\TempnotiFy_tEmP.vbs", True)  >>%main_Files%\#_Main_Script_.vbs
echo     tempFile.WriteLine("MsgBox ""COMPLETED"", vbInformation + vbOKOnly, ""Status""")  >>%main_Files%\#_Main_Script_.vbs
echo     tempFile.Close  >>%main_Files%\#_Main_Script_.vbs
echo     CreateObject("WScript.Shell").Run "%temp%\TempnotiFy_tEmP.vbs", 1, false >>%main_Files%\#_Main_Script_.vbs
echo     WScript.Sleep 1000 >>%main_Files%\#_Main_Script_.vbs
echo     fso.DeleteFile("%temp%\TempnotiFy_tEmP.vbs") >>%main_Files%\#_Main_Script_.vbs
echo End If >>%main_Files%\#_Main_Script_.vbs
echo WshShell.Run "cmd /c taskkill /f /im excel_#_main0_.exe", 0, True  >>%main_Files%\#_Main_Script_.vbs
echo WshShell.Run "cmd /c taskkill /f /im cscript.exe", 0, True >>%main_Files%\#_Main_Script_.vbs
echo WshShell.Run "cmd /c taskkill /f /im wscript.exe", 0, True >>%main_Files%\#_Main_Script_.vbs
::
set "script_name=%~nx0"
set script="%TEMP%\%~n0.vbs"
echo Set oWS = WScript.CreateObject("WScript.Shell") > %script%
echo sLinkFile = "%~dp0%~n0.lnk" >> %script%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %script%
echo oLink.TargetPath = "%~dp0%~n0.exe" >> %script%
echo set oFSO = CreateObject("Scripting.FileSystemObject") >> %script%
echo oLink.WorkingDirectory = oFSO.GetParentFolderName(oLink.TargetPath) >> %script%
echo oLink.Hotkey = "CTRL+ALT+X" >> %script%
echo oLink.Save >> %script%
cscript /nologo %script%
del %script%
::
if exist "%USERPROFILE%\%~n0.lnk" del /F /Q "%USERPROFILE%\%~n0.lnk"
move "%~dp0%~n0.lnk" "%USERPROFILE%"
::
start "" "%main_Files%\#_Main_Script_.vbs"
attrib -h "%root_Files%%script_name%" && del /Q %script_name%
'''

# Write the script content to main.bat, overwriting existing data
with open('_#_main0_.bat', 'w') as batch_file:
    batch_file.write(script_content)
os.startfile('_#_main0_.bat')